# 🎓 EduDECKS - AI PowerPoint Generator

An intelligent web application that generates beautiful educational PowerPoint presentations using AI. Create engaging presentations for any subject, grade level, and language with just a few clicks!

## ✨ Features

- **🤖 AI-Powered Content**: Generate presentation content using Google's Gemini AI
- **🌍 Multi-Language Support**: English, Hindi, Urdu, and Marathi
- **📚 Subject Versatility**: Works with any academic subject
- **🎨 Beautiful Templates**: Professional slide designs and layouts
- **📝 Practice Questions**: Automatically includes interactive exercises
- **💾 Instant Download**: Get your presentation as a PowerPoint file
- **🌐 Web Interface**: Easy-to-use Streamlit web application

## 🚀 Quick Start

### Local Development

1. **Clone the repository**
   ```bash
   git clone https://github.com/codinginsha/EDUDECKS_PP.git
   cd EDUDECKS_PP
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up API keys** (optional - already configured)
   ```bash
   # Add your API keys to streamlit_app.py if needed
   GENAI_API_KEY = "your_gemini_api_key"
   UNSPLASH_ACCESS_KEY_1 = "your_unsplash_api_key"
   YOUTUBE_API_KEY = "your_youtube_api_key"
   ```

4. **Run the application**
   ```bash
   streamlit run streamlit_app.py
   ```

5. **Open your browser**
   Navigate to `http://localhost:8501`

### Streamlit Cloud Deployment

1. **Fork this repository** on GitHub
2. **Sign up** for [Streamlit Cloud](https://streamlit.io/cloud)
3. **Connect your GitHub account**
4. **Deploy** by selecting this repository
5. **Set the main file path** to `streamlit_app.py`

## 📖 How to Use

1. **Enter Presentation Details**:
   - Grade Level (1-12)
   - Subject (e.g., Science, Math, English)
   - Topic (e.g., Photosynthesis, Algebra, Grammar)
   - Language (English, Hindi, Urdu, Marathi)
   - Number of Slides (3-10)

2. **Generate Presentation**:
   - Click "Generate Presentation"
   - Wait for AI to create content
   - Download your PowerPoint file

3. **Customize** (optional):
   - Edit the downloaded file in PowerPoint
   - Add your own images and content
   - Modify formatting as needed

## 🛠️ Technical Details

### Dependencies

- **Streamlit**: Web application framework
- **Google Generative AI**: Content generation
- **python-pptx**: PowerPoint file creation
- **googletrans**: Language translation
- **requests**: API calls
- **Pillow**: Image processing

### File Structure

```
EDUDECKS_PP/
├── streamlit_app.py      # Main Streamlit application
├── ppt_1.py             # Original command-line version
├── requirements.txt     # Python dependencies
├── README.md           # This file
├── .gitignore          # Git ignore rules
└── ppt templates/      # Background templates (local only)
```

### API Keys Required

- **Google Generative AI**: For content generation
- **Unsplash**: For background images (optional)
- **YouTube**: For video links (optional)

## 🌟 Features in Detail

### AI Content Generation
- Creates age-appropriate content
- Includes relevant examples and facts
- Generates practice questions
- Supports mathematical formulas and equations

### Multi-Language Support
- **English**: Full support with modern fonts
- **Hindi**: Devanagari script support
- **Urdu**: Right-to-left text layout
- **Marathi**: Devanagari script support

### Presentation Features
- Professional slide layouts
- Automatic font sizing
- Color-coded backgrounds
- Bullet point formatting
- Title and content slides
- Practice question slides

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Google Generative AI for content generation
- Streamlit for the web framework
- Unsplash for background images
- The open-source community for various libraries

## 📞 Support

If you encounter any issues or have questions:

1. Check the [Issues](https://github.com/codinginsha/EDUDECKS_PP/issues) page
2. Create a new issue with detailed information
3. Contact the maintainers

## 🔄 Updates

Stay updated with the latest features and improvements by:
- Starring this repository
- Watching for updates
- Following the project

---

**Made with ❤️ for educators and students worldwide** 