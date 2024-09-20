import streamlit as st
from docx import Document
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import requests
import os

# Set the page configuration for the Streamlit app
st.set_page_config(
    page_title="Feedback Form", 
    page_icon="https://lirp.cdn-website.com/d8120025/dms3rep/multi/opt/social-image-88w.png", 
    layout="centered"
)

# Function to generate a unique ID using the UUID API
def generate_unique_id():
    try:
        response = requests.get("https://www.uuidtools.com/api/generate/v1")
        if response.status_code == 200:
            unique_id = response.json()[0]
            return unique_id
        else:
            st.warning("Error generating unique ID, fallback to internal ID.")
            return "fallback_id"  # Use a fallback ID in case of API failure
    except Exception as e:
        st.warning(f"Error generating unique ID: {e}, using fallback.")
        return "fallback_id"  # Fallback to a safe default ID

# Function to replace placeholders in the Word document
def replace_placeholder(paragraphs, placeholder, value):
    """ Replace a placeholder with a value in the Word document. """
    placeholder_with_brackets = f'[{placeholder}]'
    for paragraph in paragraphs:
        if placeholder_with_brackets in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder_with_brackets, value)

# Function to mark selected options with an 'X' in the Word document
def mark_selected_option(paragraphs, placeholder_dict):
    """ Mark the selected option with an 'X' next to it in the Word document.
        placeholder_dict: A dictionary where keys are placeholders (like p2, p3, etc.)
        and values are booleans indicating whether they are selected or not.
    """
    for paragraph in paragraphs:
        for placeholder, is_selected in placeholder_dict.items():
            placeholder_with_brackets = f'[{placeholder}]'
            if placeholder_with_brackets in paragraph.text:
                if is_selected:
                    paragraph.text = paragraph.text.replace(placeholder_with_brackets, '[X]')
                else:
                    paragraph.text = paragraph.text.replace(placeholder_with_brackets, '[ ]')

# Function to populate the Word document with form data
def populate_document(data, template_path, save_directory="/"):
    try:
        # Ensure that the save directory exists
        if not os.path.exists(save_directory):
            os.makedirs(save_directory, exist_ok=True)
        
        # Generate a unique ID for the filled document
        unique_id = generate_unique_id()

        # Load the Word document template
        doc = Document(template_path)
        paragraphs = doc.paragraphs

        # Replace course name for placeholder p1
        replace_placeholder(paragraphs, 'p1', data['course_name'])

        # Replace placeholders for questions with options
        mark_selected_option(paragraphs, {
            'p2': data['course_selection_feedback'] == "Very Satisfied",
            'p3': data['course_selection_feedback'] == "Satisfied",
            'p4': data['course_selection_feedback'] == "Neutral",
            'p5': data['course_selection_feedback'] == "Unsatisfied",
            'p6': data['course_selection_feedback'] == "Very Unsatisfied"
        })

        mark_selected_option(paragraphs, {
            'p7': data['course_info_clarity'] == "Yes",
            'p8': data['course_info_clarity'] == "No",
            'p9': data['course_info_clarity'] == "Somewhat"
        })

        replace_placeholder(paragraphs, 'p10', data['course_selection_suggestions'])

        mark_selected_option(paragraphs, {
            'p11': data['course_guidance_rating'] == "Excellent",
            'p12': data['course_guidance_rating'] == "Good",
            'p13': data['course_guidance_rating'] == "Fair",
            'p14': data['course_guidance_rating'] == "Poor"
        })

        mark_selected_option(paragraphs, {
            'p15': data['course_delivery_satisfaction'] == "Very Satisfied",
            'p16': data['course_delivery_satisfaction'] == "Satisfied",
            'p17': data['course_delivery_satisfaction'] == "Neutral",
            'p18': data['course_delivery_satisfaction'] == "Unsatisfied",
            'p19': data['course_delivery_satisfaction'] == "Very Unsatisfied"
        })

        mark_selected_option(paragraphs, {
            'p20': data['course_content_relevance'] == "Highly Relevant",
            'p21': data['course_content_relevance'] == "Relevant",
            'p22': data['course_content_relevance'] == "Somewhat Relevant",
            'p23': data['course_content_relevance'] == "Not Relevant"
        })

        replace_placeholder(paragraphs, 'p24', data['course_guidance_suggestions'])

        # Section 4 - Job Guidance and Application Support Feedback
        # 4.1 Job guidance satisfaction
        mark_selected_option(paragraphs, {
            'p25': data['job_guidance_satisfaction'] == "Very Satisfied",
            'p26': data['job_guidance_satisfaction'] == "Satisfied",
            'p27': data['job_guidance_satisfaction'] == "Neutral",
            'p28': data['job_guidance_satisfaction'] == "Unsatisfied",
            'p29': data['job_guidance_satisfaction'] == "Very Unsatisfied"
        })

        # 4.2 Support for job application submissions
        mark_selected_option(paragraphs, {
            'p30': data['job_application_helpfulness'] == "Extremely Helpful",
            'p31': data['job_application_helpfulness'] == "Very Helpful",
            'p32': data['job_application_helpfulness'] == "Moderately Helpful",
            'p33': data['job_application_helpfulness'] == "Slightly Helpful",
            'p34': data['job_application_helpfulness'] == "Not Helpful"
        })

        # 4.3 Interview preparation support - ensure only one is marked
        if data['interview_preparation_support'] == "Yes":
            mark_selected_option(paragraphs, {'p35': True, 'p36': False, 'p37': False})
        elif data['interview_preparation_support'] == "No":
            mark_selected_option(paragraphs, {'p35': False, 'p36': True, 'p37': False})
        else:  # "Somewhat"
            mark_selected_option(paragraphs, {'p35': False, 'p36': False, 'p37': True})

        # 4.4 Suggestions for improving job guidance or application support
        replace_placeholder(paragraphs, 'p38', data['job_guidance_suggestions'])

        replace_placeholder(paragraphs, 'p39', data['most_helpful_service'])
        replace_placeholder(paragraphs, 'p40', data['areas_for_improvement'])
        replace_placeholder(paragraphs, 'p41', data['other_comments'])

        # Save the filled document in the current directory with a unique name
        filled_doc_path = f"Filled_Student_Feedback_Form_{unique_id}.docx"
        doc.save(filled_doc_path)

        return filled_doc_path

    except Exception as e:
        st.error(f"Error processing the document: {e}")
        return None

# Function to send the document via Outlook
def send_email(file_path):
    try:
        sender_email = st.secrets["sender_email"]
        password = st.secrets["sender_password"]
        receiver_email=sender_email
        smtp_server = "smtp.office365.com"
        smtp_port = 587

        # Check if the file exists before attempting to send
        if not os.path.exists(file_path):
            st.warning("File not found. Skipping email sending.")
            return

        # Create a multipart message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = "Student Feedback Form Submission"

        # Email body
        body = "Please find the attached filled feedback form."
        msg.attach(MIMEText(body, 'plain'))

        # Attach the document
        with open(file_path, "rb") as attachment:
            part = MIMEApplication(attachment.read(), _subtype="docx")
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
            msg.attach(part)

        # Setup the server and send the email
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure the connection
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()

        st.success(f"Feedback form submitted and sent to {receiver_email}.")

    except smtplib.SMTPException as smtp_error:
        st.error(f"SMTP error occurred: {smtp_error}")

    except FileNotFoundError as file_error:
        st.error(f"Error with file handling: {file_error}")

    except Exception as e:
        st.error(f"An error occurred while sending the email: {e}")

# Streamlit form for user input with `key` to store values in session state
st.title("Student Feedback Form")

# Feedback form fields with `key` for session state
st.subheader("1. Course Name")
course_name = st.text_input("Course Name", key="course_name")

st.subheader("2. Course Selection Feedback")
course_selection_feedback = st.selectbox(
    "How satisfied are you with the course selection process?", 
    ["Very Satisfied", "Satisfied", "Neutral", "Unsatisfied", "Very Unsatisfied"], 
    key="course_selection_feedback"
)
course_info_clarity = st.selectbox(
    "Was the information about the courses clear?", 
    ["Yes", "No", "Somewhat"], 
    key="course_info_clarity"
)
course_selection_suggestions = st.text_area("Suggestions to improve the course selection process", key="course_selection_suggestions")

st.subheader("3. Course Guidance and Delivery Feedback")
course_guidance_rating = st.selectbox(
    "Rate the guidance when selecting the right course:", 
    ["Excellent", "Good", "Fair", "Poor"], 
    key="course_guidance_rating"
)
course_delivery_satisfaction = st.selectbox(
    "How satisfied are you with the course delivery?", 
    ["Very Satisfied", "Satisfied", "Neutral", "Unsatisfied", "Very Unsatisfied"], 
    key="course_delivery_satisfaction"
)
course_content_relevance = st.selectbox(
    "Was the course content relevant to your career goals?", 
    ["Highly Relevant", "Relevant", "Somewhat Relevant", "Not Relevant"], 
    key="course_content_relevance"
)
course_guidance_suggestions = st.text_area("Suggestions to improve course guidance or delivery", key="course_guidance_suggestions")

st.subheader("How satisfied are you with the job guidance services?")
job_guidance_satisfaction = st.selectbox(
    "How satisfied are you with the job guidance services?", 
    ["Very Satisfied", "Satisfied", "Neutral", "Unsatisfied", "Very Unsatisfied"], 
    key="job_guidance_satisfaction"
)
job_application_helpfulness = st.selectbox(
    "How helpful was the job application support?", 
    ["Extremely Helpful", "Very Helpful", "Moderately Helpful", "Slightly Helpful", "Not Helpful"], 
    key="job_application_helpfulness"
)
interview_preparation_support = st.selectbox(
    "Did you receive adequate support for interview preparation?", 
    ["Yes", "No", "Somewhat"], 
    key="interview_preparation_support"
)
job_guidance_suggestions = st.text_area("Suggestions to improve job guidance or application support", key="job_guidance_suggestions")

st.subheader("5. Additional Feedback")
most_helpful_service = st.text_area("What did you find most helpful about our services?", key="most_helpful_service")
areas_for_improvement = st.text_area("What areas need improvement?", key="areas_for_improvement")
other_comments = st.text_area("Any other comments or suggestions?", key="other_comments")


# Submit button
if st.button("Submit", key="submit_button"):
    try:
        # Validate form inputs
        if not course_name :
            st.error("Please fill in all required fields, including course name and recipient email.")
        else:
            # Collect form data from session state
            form_data = {
                'course_name': st.session_state.course_name,
                'course_selection_feedback': st.session_state.course_selection_feedback,
                'course_info_clarity': st.session_state.course_info_clarity,
                'course_selection_suggestions': st.session_state.course_selection_suggestions,
                'course_guidance_rating': st.session_state.course_guidance_rating,
                'course_delivery_satisfaction': st.session_state.course_delivery_satisfaction,
                'course_content_relevance': st.session_state.course_content_relevance,
                'course_guidance_suggestions': st.session_state.course_guidance_suggestions,
                'job_guidance_satisfaction': st.session_state.job_guidance_satisfaction,
                'job_application_helpfulness': st.session_state.job_application_helpfulness,
                'interview_preparation_support': st.session_state.interview_preparation_support,
                'job_guidance_suggestions': st.session_state.job_guidance_suggestions,
                'most_helpful_service': st.session_state.most_helpful_service,
                'areas_for_improvement': st.session_state.areas_for_improvement,
                'other_comments': st.session_state.other_comments,
            }

            # Path to the Word template document
            template_path = "resource/ph_feedback_form.docx"  # Correct path

            # Populate the document
            filled_doc_path = populate_document(form_data, template_path)

            # Send the document via email
            if filled_doc_path:
                send_email(filled_doc_path)

    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
