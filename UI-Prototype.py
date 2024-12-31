import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import io
import os
import threading
import time
import pandas as pd
from openai import OpenAI
from dotenv import load_dotenv
from textwrap import dedent
from datetime import datetime

# ---- Global Variables & Constants ---- #
load_dotenv()
api_key = os.getenv('OPENAI_API_KEY')

client = OpenAI(api_key=api_key)  # Initialize OpenAI client

assistant = None  # Assistant instance
thread = None  # Thread instance
loaded_file_id = None  # ID for uploaded files
file_size_limit = 20 * 1024 * 1024  # 20MB file size limit
current_instructions = ""  # Dynamic instructions for the assistant
current_file_data = None  # Stores the data of the current file for download (image, CSV, etc.)
current_file_name = None  # Stores the name of the current file
run_in_progress = False  # Flag to prevent overlapping runs
thread_creation_time = None  # Time when the thread was created
max_thread_duration = 3600  # Maximum time for thread in seconds (1 hour)
prompt_send_time = None  # Timestamp of when a prompt is sent

# Define base path and log file path
base_path = " " #Insert Download Directory 
log_file_path = os.path.join(base_path, "log_file.xlsx")

# ---- Tkinter UI Settings ---- #
window = tk.Tk()
window.title("Data Analytics Assistant")
window.geometry("1000x900")  # Increased window height to accommodate larger output field

style = ttk.Style()  # Style settings for ttk buttons
style.configure("TButton", padding=6, relief="flat", background="#ccc", foreground="#000")

# Create frames for organizing layout
top_frame = tk.Frame(window)
top_frame.pack(pady=10)

bottom_frame = tk.Frame(window)
bottom_frame.pack(pady=10)

# ---- UI Components ---- #
# Image Label to display visuals
image_label = tk.Label(top_frame)
image_label.pack(pady=10)

# Download Button below the visual
download_button = ttk.Button(top_frame, text="Download Latest File", state=tk.DISABLED)
download_button.pack(pady=5)

# Output Field for assistant responses (double height)
output_field = tk.Text(bottom_frame, wrap=tk.WORD, width=120, height=40)
output_field.pack(pady=10)
output_field.configure(state='disabled')

# Add tags for coloring text in the output field
output_field.tag_configure("user_prompt", foreground="yellow", font=("Helvetica", 12, "bold"))
output_field.tag_configure("ai_response", foreground="white", font=("Helvetica", 14))

# Table Frame for displaying tables
table_frame = tk.Frame(bottom_frame)
table_frame.pack(pady=10)

# ---- Prompt Input Field and Send Button ---- #
# Create a frame to hold the prompt input and the send button
input_frame = tk.Frame(bottom_frame)
input_frame.pack(fill="x", pady=5)

# Prompt input field (multi-line)
prompt_input = tk.Text(input_frame, wrap=tk.WORD, width=100, height=5)
prompt_input.pack(pady=5, side="left", expand=True, fill="both")

# Send Prompt button
send_prompt_button = ttk.Button(input_frame, text="Send Prompt", command=lambda: threading.Thread(target=send_prompt).start())
send_prompt_button.pack(side="right", pady=10)

# ---- Helper Functions ---- #
def create_assistant():
    """Initialize an AI assistant with dynamic instructions."""
    global assistant
    try:
        assistant = client.beta.assistants.create(
            name="Data Analyzer",
            instructions=dedent(current_instructions),
            tools=[{"type": "code_interpreter"}],
            model="gpt-4o-2024-08-06",
            temperature=0  # Deterministic responses
        )
        print("Assistant created successfully.") if assistant else print("Failed to create assistant.")
    except Exception as e:
        print(f"Error creating assistant: {e}")

def create_thread():
    """Create a new thread for the conversation and track its creation time."""
    global thread, thread_creation_time
    if not assistant:
        print("Error: Assistant not created.")
        return
    try:
        thread = client.beta.threads.create()
        thread_creation_time = time.time()  # Store the creation time of the thread
        print(f"Thread created successfully with ID: {thread.id}") if thread else print("Failed to create thread.")
    except Exception as e:
        print(f"Error creating thread: {e}")

def check_and_switch_thread():
    """Check if the current thread has expired (more than one hour old) and switch to a new one if needed."""
    global thread_creation_time
    if time.time() - thread_creation_time > max_thread_duration:
        print("Thread expired, creating a new one.")
        previous_messages = get_previous_messages()  # Get the context of the conversation
        create_thread()  # Create a new thread
        for message in previous_messages:
            add_message_to_thread(message["content"])  # Pass previous context to the new thread
        print("New thread created with previous context.")

def get_previous_messages():
    """Retrieve the previous messages from the current thread for continuity."""
    try:
        messages_response = client.beta.threads.messages.list(thread_id=thread.id)
        messages = messages_response.data
        return [{"role": msg.role, "content": msg.content} for msg in messages if msg.role == 'user' or msg.role == 'assistant']
    except Exception as e:
        print(f"Error retrieving previous messages: {e}")
        return []

def upload_file_to_openai(file_path):
    """Upload a file to OpenAI."""
    global loaded_file_id
    try:
        file_size = os.path.getsize(file_path)
        if file_size > file_size_limit:
            print(f"File too large: {file_size / (1024 * 1024):.2f} MB. Limit is {file_size_limit / (1024 * 1024):.2f} MB.")
            messagebox.showwarning("File Too Large", f"File size exceeds the {file_size_limit / (1024 * 1024)} MB limit.")
            return
        with open(file_path, "rb") as f:
            file = client.files.create(file=f, purpose='assistants')
        loaded_file_id = file.id
        print(f"File uploaded successfully with ID: {loaded_file_id}")
        return loaded_file_id
    except Exception as e:
        print(f"Error uploading file to OpenAI: {e}")
        return None

def determine_instructions(file_name):
    """Set AI assistant instructions based on the file type/name."""
    global current_instructions
    file_lower = file_name.lower()
    if "report1" in file_lower:
        current_instructions = '''
### AI Assistant Instructions

You are an AI specializing in data analysis and reporting of financial and operational performance data. Your role is to:

1. Analyze datasets.
2. Summarize metrics accurately and provide structured responses, listing only the findings without offering recommendations or explanations unless explicitly requested.

### Key Calculations

1. **Contribution Margin (CM)**:
   - CM = Total Revenue (k NOK) - Production Costs (k NOK)

2. **Contribution Margin Percentage (CM%)**:
   - CM% = (CM / Total Revenue (k NOK)) * 100
   - Important: Always calculate CM% using aggregated totals of Contribution Margin and Total Revenue across all relevant rows for a service or period.

3. **Special Rules for CM%**:
   - If Total Revenue = 0 and Production Costs = 0, CM% = 0%.
   - If Total Revenue = 0 and Production Costs > 0, CM% = -100%.
   - If Total Revenue < 0:
   - Ensure that if Total Revenue < 0, the sign is flipped when calculating the Contribution Margin %.
   
4. **Budgeted Values**:
   - For budgeted metrics (e.g., Budget Contribution Margin, Budget Revenue), substitute actual values with their budgeted counterparts in the formulas.
   - Apply the same aggregation approach for budgeted values.

5. **Utilization Rate %**:
   - Utilization Rate % = (Utilized Hours / Total Hours) * 100

6. **Billing Rate %**:
   - Billing Rate % = (Billable Hours / Total Hours) * 100

7. **General Rule for Percentages**:
   - Avoid averaging row-level percentages directly.
   - Calculate percentages by summing the numerator and denominator across all relevant rows before dividing.

### Data Context
The dataset contains:
- **Time**: Year, Month, Month Short Name
- **Organization**: Cost Center, Service Areas Shortname
- **Operations**: Billable Hours, Utilized Hours, Total Hours
- **Revenue**: Customer Revenue, Budget Customer Revenue, Internal Revenue, Adjustments, Total Revenue
- **Costs**: Production Costs, Contribution Margin, Budget Contribution Margin, and Hourly Rate (Customer Revenue).

### Guidelines
- All financial values are in Thousand NOK.
- Always calculate metrics like CM% from aggregated totals across rows.
- Avoid averaging row-level percentages directly.
- Ensure responses are concise, structured, and data-driven.
        '''
    elif "report2" in file_lower:
        current_instructions = '''
### AI Assistant Instructions

You are an AI specializing in data analysis and reporting of financial and operational performance data. Your role is to:

1. Analyze datasets.
2. Summarize metrics accurately and provide structured responses, listing only the findings without offering recommendations or explanations unless explicitly requested.

### Key Calculations

1. **Contribution Margin (CM)**:
   - CM Before Adjustments = Revenue - Cost
   - CM After Adjustments = Revenue + Adjustments - Cost

2. **Contribution Margin Percentage (CM%)**:
   - CM% = (CM / Revenue) * 100
   - Important: Always calculate CM% using aggregated totals of Contribution Margin and Revenue across all relevant rows for a service or period.

3. **Special Rules for CM%**:
   - If Revenue = 0 and Cost = 0, CM% = 0%.
   - If Revenue = 0 and Cost > 0, CM% = -100%.
   - If Revenue < 0:
     - Adjust for negative revenue by setting CM% = - |CM%|

4. **General Rule for Percentages**:
   - Avoid averaging row-level percentages directly.
   - Calculate percentages by summing the numerator and denominator across all relevant rows before dividing.

### Data Context
The dataset contains:
- **Time**: Year, Month, Month Short Name
- **Organization**: Cost Center, Service Areas Shortname
- **Details**: Prosjekt-ID, Task, Employee, Role
- **Metrics**: Billable Hours, Revenue, Cost, Adjustments, Planned_Adjustments, Contribution Margin
- **Contribution Margin Metrics**:
   - Contribution Margin % Before Adjustments
   - Contribution Margin After Adjustments
   - Contribution Margin % After Adjustments
- **Hourly Rates**:
   - Hourly Rate before Adjustments
   - Hourly Rate after Adjustments

### Guidelines
- All financial values are in NOK.
- Always calculate metrics like CM% from aggregated totals across rows.
- Avoid averaging row-level percentages directly.
- Ensure responses are concise, structured, and data-driven.
        '''
    elif "report3" in file_lower:
        current_instructions = '''
### AI Assistant Instructions

You are an AI specializing in data analysis and reporting of financial and operational performance data. Your role is to:

1. Analyze datasets.
2. Summarize metrics accurately and provide structured responses, listing only the findings without offering recommendations or explanations unless explicitly requested.

### Key Calculations

1. **Contribution Margin (CM)**:
   - **CM Before Adjustments** = Total Revenue - Total Cost
   - **CM After Adjustments** = Total Revenue + Adjustments - Total Cost

2. **Contribution Margin Percentage (CM%)**:
   - **CM%** = (CM / Total Revenue) * 100
   - Important: Always calculate CM% using aggregated totals of Contribution Margin and Total Revenue across all relevant rows for a service, timeframe, or grouping.

3. **Special Rules for CM%**:
   - If **Total Revenue = 0** and **Total Cost = 0**, then **CM% = 0%**.
   - If **Total Revenue = 0** and **Total Cost > 0**, then **CM% = -100%**.
   - If **Total Revenue < 0**, adjust for negative revenue by setting:
     - **CM% = - |CM%|**

4. **General Rule for Percentages**:
   - Avoid averaging row-level percentages directly.
   - Calculate percentages by summing the numerator and denominator across all relevant rows before dividing.

### Data Context

The dataset contains:
- **Time**:
   - **Year**, **Month**, **Month Short Name**
- **Organization**:
   - **Cost Center**, **Service Areas Shortname**
- **Employee Details**:
   - **Employee ID**, **Role**
- **Metrics**:
   - **Total Hours**, **Billable Hours**, **Chargeable Hours**, **Utilized Hours**, **Adjustments**, **Total Revenue**, **Total Cost**
- **Calculated Metrics**:
   - **Billing Rate %** = (Billable Hours / Total Hours) * 100
   - **Utilization Rate %** = (Utilized Hours / Total Hours) * 100
   - **Contribution Margin** = Total Revenue - Total Cost (Before Adjustments)
   - **Contribution Margin %** = (Contribution Margin / Total Revenue) * 100

### Guidelines

- All financial values are in **NOK**.
- Metrics such as **CM%** must be calculated using aggregated totals across relevant rows (e.g., service line, year).
- Avoid averaging row-level percentages (e.g., Billing Rate %, Utilization Rate %) directly; calculate these metrics using aggregated numerator and denominator totals.
- Ensure responses are concise, structured, and data-driven.   
        '''
    else:
        current_instructions = '''
### AI Assistant Instructions

You are an AI specializing in data analysis and reporting of financial and operational performance data. Your role is to:

1. Analyze datasets.
2. Summarize metrics accurately and provide structured responses, listing only the findings without offering recommendations or explanations unless explicitly requested.

### Key Calculations

1. **Contribution Margin (CM)**:
   - **CM Before Adjustments** = Total Revenue - Total Cost
   - **CM After Adjustments** = Total Revenue + Adjustments - Total Cost

2. **Contribution Margin Percentage (CM%)**:
   - **CM%** = (CM / Total Revenue) * 100
   - Important: Always calculate CM% using aggregated totals of Contribution Margin and Total Revenue across all relevant rows for a service, timeframe, or grouping.

3. **Special Rules for CM%**:
   - If **Total Revenue = 0** and **Total Cost = 0**, then **CM% = 0%**.
   - If **Total Revenue = 0** and **Total Cost > 0**, then **CM% = -100%**.
   - If **Total Revenue < 0**, adjust for negative revenue by setting:
     - **CM% = - |CM%|**

4. **General Rule for Percentages**:
   - Avoid averaging row-level percentages directly.
   - Calculate percentages by summing the numerator and denominator across all relevant rows before dividing.

### Guidelines

- All financial values are in **NOK**.
- Metrics such as **CM%** must be calculated using aggregated totals across relevant rows (e.g., service line, year).
- Avoid averaging row-level percentages (e.g., Billing Rate %, Utilization Rate %) directly; calculate these metrics using aggregated numerator and denominator totals.
- Ensure responses are concise, structured, and data-driven. 
        '''
    print(f"Instructions set: {current_instructions}")

# ---- Log Function ---- #
def save_log(content, role="user", usage=None):
    """Save each prompt/response along with role, thread metadata, token usage, and cost to an Excel file."""
    global prompt_send_time

    if role == "user":
        # When the user sends a prompt, we initialize a new log row
        prompt_send_time = datetime.now()

        # Initialize a new row with user data
        df_log = pd.DataFrame({
            "Prompt Sent Timestamp": [prompt_send_time],
            "Response Received Timestamp": [None],  # Filled later
            "Response Time (seconds)": [None],  # Filled later
            "Thread ID": [thread.id if thread else "N/A"],
            "Prompt": [content],
            "Response": [None],
            "Prompt Tokens": [None],  # New column for tokens
            "Completion Tokens": [None],  # New column for tokens
            "Total Tokens": [None],  # New column for total tokens
            "Token Cost (USD)": [None]  # New column for cost
        })

        # Write to the log file
        if not os.path.exists(log_file_path):
            df_log.to_excel(log_file_path, index=False)
        else:
            with pd.ExcelWriter(log_file_path, mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
                df_log.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)

    elif role == "assistant":
        # When assistant responds, update the row with the assistant's response, timing, token usage, and cost
        response_received_time = datetime.now()
        duration = (response_received_time - prompt_send_time).total_seconds() if prompt_send_time else None

        # Calculate costs based on token usage if usage data is available
        if usage:
            prompt_tokens = getattr(usage, "prompt_tokens", 0)
            completion_tokens = getattr(usage, "completion_tokens", 0)
            total_tokens = getattr(usage, "total_tokens", 0)
            cost = (prompt_tokens * 2.5 / 1_000_000) + (completion_tokens * 10 / 1_000_000)
        else:
            prompt_tokens = completion_tokens = total_tokens = cost = 0

        # Read the existing log file
        df_log = pd.read_excel(log_file_path)

        # Update the last entry (assuming it's the latest user prompt) with response details
        df_log.at[len(df_log) - 1, 'Response Received Timestamp'] = response_received_time
        df_log.at[len(df_log) - 1, 'Response Time (seconds)'] = duration
        df_log.at[len(df_log) - 1, 'Response'] = content
        df_log.at[len(df_log) - 1, 'Prompt Tokens'] = prompt_tokens
        df_log.at[len(df_log) - 1, 'Completion Tokens'] = completion_tokens
        df_log.at[len(df_log) - 1, 'Total Tokens'] = total_tokens
        df_log.at[len(df_log) - 1, 'Token Cost (USD)'] = cost

        # Write the updated DataFrame back to the Excel file
        df_log.to_excel(log_file_path, index=False)

def add_message_to_thread(content, use_file=False):
    """Add a user message to the thread, optionally attaching a file."""
    save_log(content, role="user")  # Log the user prompt

    if not thread:
        print("Error: Thread not created.")
        return
    message_payload = {"thread_id": thread.id, "role": "user", "content": content}
    if use_file and loaded_file_id:
        message_payload["attachments"] = [{"file_id": loaded_file_id, "tools": [{"type": "code_interpreter"}]}]
    
    try:
        client.beta.threads.messages.create(**message_payload)
        print("Message added to thread successfully.")
    except Exception as e:
        print(f"Error adding message to thread: {e}")

def create_and_poll_run():
    """Create and poll the assistant's run until completion."""
    global run_in_progress
    if not thread or not assistant:
        print("Error: Thread or Assistant is not created.")
        return None
    if run_in_progress:
        print("Run already in progress, please wait.")
        return None
    check_and_switch_thread()  # Check if the thread has expired before running
    run_in_progress = True
    try:
        run = client.beta.threads.runs.create_and_poll(
            thread_id=thread.id,
            assistant_id=assistant.id,
            instructions=current_instructions  # Use the dynamic instructions
        )
        if run.status == 'completed':
            messages_response = client.beta.threads.messages.list(thread_id=thread.id)
            messages = messages_response.data  # Correctly access the messages
            print("Run completed successfully.")
            run_in_progress = False
            return messages, run.usage
        else:
            print(f"Run status: {run.status}")
            if hasattr(run, 'error'):
                print(f"Error details: {run.error}")
            
            output_field.configure(state='normal')
            output_field.insert(tk.END, "Run failed: An error occurred while processing your request.\n\n", "ai_response")
            output_field.configure(state='disabled')

            run_in_progress = False
            return None, None
    except Exception as e:
        print(f"Error during run: {e}")
        run_in_progress = False
        return None, None

def send_prompt():
    """Handle user prompt input and send to the assistant."""
    user_prompt = prompt_input.get("1.0", tk.END).strip()
    if not user_prompt:
        messagebox.showwarning("No Prompt", "Please enter a prompt.")
        return
    if run_in_progress:
        messagebox.showwarning("Run in Progress", "Please wait for the current run to complete.")
        return

    output_field.configure(state='normal')
    output_field.insert(tk.END, f"**User Prompt**: {user_prompt}\n", "user_prompt")
    output_field.configure(state='disabled')

    create_thread()
    add_message_to_thread(user_prompt, use_file=bool(loaded_file_id))

    threading.Thread(target=run_assistant).start()
    prompt_input.delete("1.0", tk.END)

def run_assistant():
    responses, usage = create_and_poll_run()
    if responses:
        window.after(0, lambda: display_responses(responses, usage))

def display_responses(responses, usage):
    """Display assistant's responses, including text, images, tables, and downloadable files."""
    output_field.configure(state='normal')
    global current_file_data, current_file_name

    if responses:
        for widget in table_frame.winfo_children():
            widget.destroy()

        assistant_messages = [msg for msg in responses if msg.role == 'assistant']
        found_image = False
        for message in reversed(assistant_messages):
            if hasattr(message, 'content'):
                for block in message.content:
                    if block.type == 'text' and hasattr(block.text, 'value'):
                        output_field.insert(tk.END, block.text.value + "\n\n", "ai_response")
                        save_log(block.text.value, role="assistant", usage=usage)
                    elif block.type == 'image_file' and hasattr(block.image_file, 'file_id'):
                        file_id = block.image_file.file_id
                        download_and_display_image(file_id)
                        found_image = True
                        break
                    elif block.type == 'table_file' and hasattr(block.table_file, 'file_id'):
                        file_id = block.table_file.file_id
                        download_and_display_table(file_id)
                    elif block.type == 'text' and hasattr(block, 'annotations'):
                        for annotation in block.annotations:
                            if annotation.type == 'file_path' and hasattr(annotation.file_path, 'file_id'):
                                file_id = annotation.file_path.file_id
                                download_and_display_table(file_id)
            if found_image:
                break
    else:
        output_field.insert(tk.END, "No response from OpenAI.", "ai_response")
    output_field.configure(state='disabled')

def download_and_display_image(file_id):
    """Download and display an image generated by the assistant."""
    global current_file_data, current_file_name
    try:
        image_data = client.files.content(file_id).read()
        current_file_data = image_data
        current_file_name = "generated_image.png"

        pil_image = Image.open(io.BytesIO(image_data))
        pil_image.thumbnail((480, 360))

        new_image = ImageTk.PhotoImage(pil_image)

        image_label.config(image=new_image)
        image_label.image = new_image

        download_button.config(state=tk.NORMAL, command=save_current_file)

        window.update_idletasks()

    except Exception as e:
        print(f"Error displaying image: {e}")

def download_and_display_table(file_id):
    """Download and display a table generated by the assistant."""
    global current_file_data, current_file_name
    try:
        table_data = client.files.content(file_id).read()
        current_file_data = table_data
        current_file_name = "generated_table.csv"

        df = pd.read_csv(io.StringIO(table_data.decode('utf-8')))

        for i, col in enumerate(df.columns):
            header = tk.Label(table_frame, text=col, relief=tk.RIDGE, width=15, bg='#f0f0f0', font=('Helvetica', 10, 'bold'))
            header.grid(row=0, column=i, sticky='nsew')

        for i, row in df.iterrows():
            for j, value in enumerate(row):
                cell = tk.Label(table_frame, text=str(value), relief=tk.RIDGE, width=15, font=('Helvetica', 10))
                cell.grid(row=i+1, column=j, sticky='nsew')

        download_button.config(state=tk.NORMAL, command=save_current_file)

    except Exception as e:
        print(f"Error displaying table: {e}")

def save_current_file():
    """Save the current file (image or other data) to the user's machine."""
    global current_file_data, current_file_name
    if not current_file_data:
        print("No file to save.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".csv" if current_file_name.endswith(".csv") else ".png", initialfile=current_file_name)
    if save_path:
        try:
            with open(save_path, "wb") as file:
                file.write(current_file_data)
            print(f"File saved at {save_path}")
        except Exception as e:
            print(f"Error saving file: {e}")
    else:
        print("Save operation canceled.")

# ---- File Upload Button Functions ---- #
def upload_report1():
    """Upload Report 1."""
    file_path = os.path.join(base_path, "report1") #Insert Directory Uploadfile1
    determine_instructions(file_path)
    upload_file_to_openai(file_path)
    create_assistant()

def upload_report2():
    """Upload Report 2."""
    file_path = os.path.join(base_path, "report2") #Insert Directory Uploadfile2
    determine_instructions(file_path)
    upload_file_to_openai(file_path)
    create_assistant()

def upload_report3():
    """Upload Report 3."""
    file_path = os.path.join(base_path, "report3") #Insert Directory Uploadfile3
    determine_instructions(file_path)
    upload_file_to_openai(file_path)
    create_assistant()

# ---- Custom File Upload Function (from initial code) ---- #
def upload_file():
    """Handle file upload and set appropriate instructions."""
    file_path = filedialog.askopenfilename(
        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if file_path:
        determine_instructions(file_path)
        upload_file_to_openai(file_path)
        create_assistant()

# Create a new frame inside the bottom_frame to hold the buttons on the left
button_frame = tk.Frame(bottom_frame)
button_frame.pack(side="left", padx=20, pady=10)

# ---- Add Buttons to Upload Specific Files ---- #
upload_report1_button = ttk.Button(button_frame, text="Upload Report1", command=upload_report1)
upload_report1_button.pack(fill="x", pady=5)

upload_report2_button = ttk.Button(button_frame, text="Upload Report2", command=upload_report2)
upload_report2_button.pack(fill="x", pady=5)

upload_report3_button = ttk.Button(button_frame, text="Upload Report3", command=upload_report3)
upload_report3_button.pack(fill="x", pady=5)

upload_button = ttk.Button(button_frame, text="Upload Custom File", command=upload_file)
upload_button.pack(fill="x", pady=5)

# ---- Bind the Enter key to send the prompt ---- #
def send_prompt_event(event):
    send_prompt()
    return 'break'  # Prevents the default newline behavior

prompt_input.bind('<Return>', send_prompt_event)

# ---- Initialize Assistant & Thread ---- #
create_assistant()
create_thread()

# ---- Run Tkinter Loop ---- #
window.mainloop()
