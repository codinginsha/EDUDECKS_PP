import google.generativeai as genai
import requests
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE
import os
import random
import re
import string
from googletrans import Translator
from googleapiclient.discovery import build
import datetime
# --- Load environment variables from .env file if present ---
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    print("(Optional) Install python-dotenv to use .env files for environment variables.") 