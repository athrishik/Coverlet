# Coverlet
CoverLet: A sleek and lightweight Streamlit web application that helps job seekers create perfectly tailored cover letters using AI assistance. This tool generates _comprehensive prompts_ for ChatGPT, Claude, or other AI models to customize cover letters based on specific job descriptions and personal information.

**✨ Features**

📝 Multiple Input Methods

Cover Letter Templates: Upload Word documents (.docx) or paste text directly
Resume Support: Upload both Word (.docx) and PDF formats with automatic text extraction
Job Description Input: Large text area for complete job posting details

🎯 Smart Language Options

Professional Language: Formal, polished tone
Match Template Language: Maintains your original writing style
Custom Instructions: Personalized tone and language preferences

🔧 Intelligent Features

One-Page Option: Optional constraint for concise cover letters
Editable AI Prompts: Fully customizable instruction templates
Document Preview: Expandable sections to review uploaded content
Real-time Prompt Generation: Dynamic updates based on your selections

📄 Output Generation

Word Document Export: Creates a complete .docx file containing:

AI prompt instructions
Resume information
Job description
Your cover letter template


One-Click Download: Timestamped filename for easy organization
AI-Ready Format: Optimized for ChatGPT, Claude, and other AI models

🎨 Modern UI/UX

Dark Theme: GitHub-inspired professional design
Clean Interface: Intuitive two-column layout
Responsive Design: Works on desktop and mobile devices
Visual Feedback: Success indicators and error handling
Professional Styling: Inter font family and consistent spacing

🛠️ Technical Stack

Frontend: Streamlit with custom CSS
Document Processing: python-docx for Word files, PyPDF2 for PDF parsing
File Handling: Robust error handling and format validation
Export: Binary document generation with proper MIME types

💡 How It Works

Upload/Paste your cover letter template
Upload/Paste your resume information
Select your preferred language style
Paste the target job description
Generate a comprehensive Word document
Download and copy the content into your preferred AI tool
Let AI customize your cover letter and remove the instructions

🎯 Perfect For

Job seekers applying to multiple positions
Career changers needing tailored messaging
Professionals who want consistent, high-quality applications
Anyone looking to leverage AI for personalized cover letters

📦 Installation
bash# Clone the repository
git clone https://github.com/yourusername/ai-cover-letter-generator.git

# Install dependencies
pip install streamlit python-docx PyPDF2

# Run the application
streamlit run app.py

🚀 Usage
Simply run the Streamlit app and follow the intuitive interface. The tool generates everything you need for AI-powered cover letter customization in one convenient document.

Keywords: Cover Letter, AI, Streamlit, Job Application, Resume, ChatGPT, Claude, Document Generation, Career Tools, Job Search
