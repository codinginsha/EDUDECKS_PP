import streamlit as st
import os
from ppt import generate_ppt_content, create_powerpoint, sanitize_filename
import tempfile

st.set_page_config(page_title="AI Slide Generator", layout="centered")
st.title("📊 AI PowerPoint Slide Generator")

grade_levels = [str(i) for i in range(1, 13)]
subjects = [
    "Science", "Math", "English", "Urdu", "Hindi", "Marathi", "Geography", "History", "Computer",
    "Physics", "Chemistry", "Biology", "Social Studies", "Islamiat", "Civics", "Economics", "General Knowledge",
    "Environmental Science", "Art", "Physical Education", "Moral Science", "Literature", "Grammar", "Poetry"
]

with st.form("ppt_form"):
    col1, col2 = st.columns(2)
    with col1:
        class_level = st.selectbox("Grade level", options=["Select grade level"] + grade_levels, index=0)
        subject = st.selectbox("Subject", options=["Select subject"] + subjects, index=0)
        topic = st.text_input("Topic", value="")
    with col2:
        language = st.selectbox("Language", ["Select language", "English", "Urdu", "Hindi", "Marathi"], index=0)
        num_slides = st.number_input("Number of slides", min_value=3, max_value=10, value=5, step=1)
    submitted = st.form_submit_button("Generate Presentation")

if submitted:
    # Validate required fields
    errors = []
    if class_level == "Select grade level":
        errors.append("Please select a grade level.")
    if subject == "Select subject":
        errors.append("Please select a subject.")
    if language == "Select language":
        errors.append("Please select a language.")
    if not topic.strip():
        errors.append("Please enter a topic.")
    if errors:
        for err in errors:
            st.error(err)
    else:
        with st.spinner("Generating slides, please wait..."):
            try:
                st.write("Step 1: Generating content...")
                slides = generate_ppt_content(class_level, subject, topic, language, num_slides)
                st.write("Step 2: Content generated.")
                if not isinstance(slides, list):
                    st.error(f"Error: {slides}")
                else:
                    import datetime
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = sanitize_filename(f"Class{class_level}_{subject}_{topic}_presentation_{timestamp}.pptx".replace(" ", ""))
                    st.write("Step 3: Creating PowerPoint file...")
                    create_powerpoint(slides, topic, class_level, subject, language)
                    st.write("Step 4: PowerPoint file created. Looking for file...")
                    generated_file = None
                    for f in os.listdir("."):
                        if f.endswith(".pptx") and f.startswith(f"Class{class_level}_{subject}_{topic}_presentation_"):
                            generated_file = f
                            break
                    if generated_file:
                        st.write(f"Step 5: Found file {generated_file}. Preparing download...")
                        with open(generated_file, "rb") as file_obj:
                            pptx_bytes = file_obj.read()
                        st.success("Presentation generated!")
                        st.download_button(
                            label="Download PPTX",
                            data=pptx_bytes,
                            file_name=generated_file,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                        os.remove(generated_file)
                    else:
                        st.error("Could not find the generated PPTX file. Please check if the file was created in the directory.")
            except Exception as e:
                st.error(f"An error occurred: {e}")
                import traceback
                st.text(traceback.format_exc()) 