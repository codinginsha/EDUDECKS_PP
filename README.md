# EduDECKS - AI PowerPoint Presentation Generator

An intelligent PowerPoint presentation generator that creates educational slides with AI-generated content, images, and videos.

## Features

- 🤖 AI-powered content generation using Google Gemini
- 🎨 Automatic background images from Unsplash
- 🎥 YouTube video integration (now links to search results for the topic and grade)
- 🌏 Multi-language support (English, Hindi, Marathi, Urdu)
- ↔️ RTL/LTR layout: Urdu and other RTL languages have flipped text/image layout and correct alignment
- 🧮 Math/numerical topics: Slides include formulas and worked examples with proper Unicode math formatting (e.g., x², π, √, fractions, etc.) in any language
- 📝 Urdu slides: Bullet points are not numbered (for natural Urdu style)
- 🖋️ Improved title cleaning and formatting for all languages (no stray English or symbols in Urdu)
- 📚 Language subjects (English, Hindi, Urdu, Marathi): If the topic is a poem or lesson, a "Difficult Words and Meanings" slide is included
- 🖼️ Dynamic layouts and font scaling for visually engaging slides
- 🆕 Each run generates a unique PowerPoint file (timestamped)

## Setup Instructions

### Prerequisites
- Python 3.7 or higher
- pip (Python package installer)

### Installation

1. **Option 1: Automatic Setup (Windows)**
   ```
   Double-click setup.bat
   ```

2. **Option 2: Manual Setup**
   ```bash
   pip install -r requirements.txt
   ```

### API Keys Required

The program uses several APIs. You'll need to update the API keys in `ppt.py`:

- **Google Gemini API**: Get from [Google AI Studio](https://makersuite.google.com/app/apikey)
- **Unsplash API**: Get from [Unsplash Developers](https://unsplash.com/developers)
- **YouTube API**: Get from [Google Cloud Console](https://console.cloud.google.com/)

## How to Run

```bash
python ppt.py
```

## Usage

1. Run the program
2. Enter the grade level (e.g., 5)
3. Enter the subject (e.g., Science, Math, History)
4. Enter the topic (e.g., Photosynthesis, Fractions, Ancient Egypt)
5. Choose language (English, Hindi, Marathi, Urdu)
6. Specify number of slides (3-10)

The program will generate a PowerPoint presentation with:
- AI-generated content
- Relevant background images
- YouTube search links for the topic and grade
- Professional formatting
- Slide transitions
- Unicode math formulas and worked examples for math/numerical topics
- RTL/LTR layout and correct bullet formatting for all languages
- "Difficult Words and Meanings" slide for language subjects (poem/lesson)
- Dynamic layouts and font scaling for consistent, engaging slides
- Unique output file for every run

## Output

The generated presentation will be saved as:
`Class{grade}_{subject}_{topic}_presentation_{timestamp}.pptx`

## Troubleshooting

- **Import errors**: Make sure all dependencies are installed using `pip install -r requirements.txt`
- **API errors**: Verify your API keys are correct and have sufficient credits
- **PowerPoint transitions**: Requires Microsoft PowerPoint to be installed for transition effects
- **Non-English fonts**: For best results in Urdu, Hindi, or Marathi, ensure you have Unicode fonts installed (e.g., Jameel Noori Nastaleeq for Urdu)
- **PermissionError when saving PPTX**: Make sure the file is not open in PowerPoint or any other app. Each run now generates a unique file, so this should not occur unless the folder is locked.

## Dependencies

- google-generativeai: AI content generation
- python-pptx: PowerPoint creation
- opencv-python: Image processing
- Pillow: Image handling
- googletrans: Translation services
- google-api-python-client: YouTube API
- comtypes: PowerPoint automation
- numpy: Numerical operations

Your dependencies are now being installed in the background. Once that's done, you can run the program with:

```bash
python ppt.py
```

**What to do next:**
1. Wait for the installation to finish (watch for any errors).
2. Open a Command Prompt in your project folder.
3. Run:  
   ```
   python ppt.py
   ```
4. Follow the prompts to generate your presentation.

If you see any errors during installation or when running the script, let me know and I’ll help you troubleshoot! 