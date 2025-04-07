import tkinter as tk
from tkinter import messagebox
import os
from datetime import datetime
import urllib.parse
import webbrowser
import re
from tkcalendar import Calendar 

email_template = """Hi {leaver_first_name} and {supervisor_first_name},

We have been notified that {leaver_first_name}'s last day will be {effective_date}.

Note that we will disable {leaver_first_name}'s account no later than 5PM (local time) on that date. Please advise us if they are planning to finish sooner so we can disable their account at the appropriate time.

{leaver_first_name}, to ensure a seamless experience for everyone, we will require that you have collected and removed all your personal items from the laptop, returned all project materials back to respective job areas, and transferred pertinent files to colleagues or supervisor. Any shared calendars or distribution groups you own should have also had their ownership transferred by then. Finally, we suggest to cancel any recurring meetings within your calendar.

{supervisor_first_name}, this would be a great opportunity to coordinate with {leaver_first_name} to hand over project related information, cancel any recurring meetings & recreate them, or acquire any other material that might not yet be saved to respective project areas.

Please factory reset your phone before returning, if you have an iPhone please follow these instructions: Go to Settings > General > Transfer or Reset iPhone > Tap Erase All Content and Settings, you will be prompted to enter your Apple ID to proceed. If you have an Android please follow these instructions: Tap Settings > Backup and reset > Factory data reset > Reset Device > Erase Everything.

Please let us know what time you are planning to come into the office on your last day to coordinate the return of your Arup equipment.

We thank you in advance for your understanding. Do not hesitate to contact us if you have any questions or concerns.

Thank you,
"""

def generate_email():
    leaver_name = leaver_name_entry.get()
    supervisor_name = supervisor_name_entry.get()
    leaver_last_day = leaver_last_day_calendar.get_date()

    if not all([leaver_name, supervisor_name, leaver_last_day]):
        messagebox.showerror("Error", "Please fill in all fields first.")
        return None, None, None, None

    try:
        effective_date = datetime.strptime(leaver_last_day, "%m/%d/%Y").strftime("%B %d, %Y")
    except ValueError:
        messagebox.showerror("Error", "Please enter date in MM/DD/YYYY format.")
        return None, None, None, None

    leaver_first_name = leaver_name.split()[0]
    supervisor_first_name = supervisor_name.split()[0]

    email = email_template.format(
        leaver_first_name=leaver_first_name,
        supervisor_first_name=supervisor_first_name,
        effective_date=effective_date,
    )

    email_text.delete(1.0, tk.END)
    email_text.insert(tk.END, email)
    return leaver_first_name, leaver_name.split()[-1], supervisor_first_name, supervisor_name.split()[-1], email

def clear_fields():
    leaver_name_entry.delete(0, tk.END)
    supervisor_name_entry.delete(0, tk.END)
    email_text.delete(1.0, tk.END)
    today = datetime.today()
    leaver_last_day_calendar.selection_set(today)
    formatted_today = today.strftime("%B %d, %Y")
    date_label.config(text=f"Selected Date: {formatted_today}")


def copy_to_clipboard():
    email_content = email_text.get(1.0, tk.END).strip()
    if email_content:
        root.clipboard_clear()
        root.clipboard_append(email_content)
        root.update()
        messagebox.showinfo("Copied", "EOT Email has been copied to clipboard.")
    else:
        messagebox.showerror("Error", "Please fill in all fields first.")

def open_in_outlook():
    result = generate_email()
    if result:
        leaver_first_name, leaver_last_name, supervisor_first_name, supervisor_last_name, email_content = result
        leaver_email = f"{leaver_first_name.lower()}.{leaver_last_name.lower()}@arup.com"
        supervisor_email = f"{supervisor_first_name.lower()}.{supervisor_last_name.lower()}@arup.com"

        subject = "Arup DT End of Term"
        body = urllib.parse.quote(email_content)
        to_emails = f"{leaver_email};{supervisor_email}"
        mailto_url = f"mailto:{to_emails}?subject={urllib.parse.quote(subject)}&body={body}"

        try:
            webbrowser.open(mailto_url)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open email client: {e}")
    else:
        messagebox.showerror("Error", "Please generate the email first.")

def validate_name(name):
    if not re.match(r"^[a-zA-Z\s'-]+$", name):
        return False
    return True

def update_date_label(event):
    selected_date_str = leaver_last_day_calendar.get_date()
    selected_date = datetime.strptime(selected_date_str, "%m/%d/%Y")
    formatted_date = selected_date.strftime("%B %d, %Y")
    date_label.config(text=f"Selected Date: {formatted_date}")

root = tk.Tk()
root.title("ARUP DT End of Term Email Generator")

tk.Label(root, text="Leaver's Name").grid(row=0, column=0, padx=10, pady=5)
leaver_name_entry = tk.Entry(root, width=40)
leaver_name_entry.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Supervisor's Name").grid(row=1, column=0, padx=10, pady=5)
supervisor_name_entry = tk.Entry(root, width=40)
supervisor_name_entry.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Leaver's Last Day").grid(row=2, column=0, padx=10, pady=5)
leaver_last_day_calendar = Calendar(root, selectmode='day', date_pattern='mm/dd/yyyy', firstweekday='sunday', showweeknumbers=False)
leaver_last_day_calendar.grid(row=2, column=1, padx=10, pady=5)
date_label = tk.Label(root, text="Selected Date: None", font=("Helvetica", 10))
date_label.grid(row=3, column=1, pady=5)
leaver_last_day_calendar.bind("<<CalendarSelected>>", update_date_label)

email_text = tk.Text(root, width=60, height=15, wrap=tk.WORD)
email_text.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

button_frame = tk.Frame(root)
button_frame.grid(row=5, column=0, columnspan=2, pady=10)

clear_button = tk.Button(button_frame, text="Clear", command=clear_fields)
clear_button.pack(side=tk.LEFT, padx=5)

generate_button = tk.Button(button_frame, text="Generate Email", command=generate_email)
generate_button.pack(side=tk.LEFT, padx=5)

copy_button = tk.Button(button_frame, text="Copy Email to Clipboard", command=copy_to_clipboard)
copy_button.pack(side=tk.LEFT, padx=5)

outlook_button = tk.Button(button_frame, text="Open in Outlook", command=open_in_outlook)
outlook_button.pack(side=tk.LEFT, padx=5)

root.mainloop()