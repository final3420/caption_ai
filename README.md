# Caption.ai — AI Marketing Caption Generator

## Overview

**Caption.ai** is a multilingual AI-powered Streamlit application for generating, comparing, analyzing, and exporting marketing captions from product images using the OpenAI GPT-4o Vision model.

The app is designed as a **Prompt Engineering Research Tool** to compare the effectiveness of:

- **Simple prompts**
- **Structured prompts**

Users can upload product images, define marketing parameters, generate captions, rate them manually or with AI, visualize results, and export experiments.

---

# Features

## ✨ AI Caption Generation
- Generate marketing captions from product images
- Uses **GPT-4o Vision**
- Supports:
  - Simple prompting
  - Structured prompting

## 🌍 Multilingual Support
- English 🇺🇸
- Arabic 🇪🇬
- RTL Arabic rendering support using:
  - `arabic_reshaper`
  - `python-bidi`

## 🖼️ Image-Based Prompting
Upload:
- PNG
- JPG
- JPEG
- WEBP

## 🧠 AI SEO Keyword Suggestions
Generate SEO keywords automatically using OpenAI.

## 📊 Caption Evaluation
Manual human rating system based on:
- Persuasiveness
- Professionalism
- Audience Fit
- Creativity

## ⚡ AI Evaluation
GPT-4o automatically:
- Scores captions
- Compares prompt quality
- Provides reasoning

## 📈 Visual Analytics
Includes:
- Radar charts
- Bar charts
- Score comparison dashboards

## 📦 Batch Processing
Generate captions for multiple products using image URLs.

## 📁 Export Options
Export:
- JSON experiment reports
- CSV rating datasets
- Batch caption CSV files

---

# Installation

## 1. Clone Repository

```bash
git clone https://github.com/yourusername/caption-ai.git
cd caption-ai
```

## 2. Create Virtual Environment

### Windows

```bash
python -m venv venv
venv\Scripts\activate
```

### Linux / macOS

```bash
python -m venv venv
source venv/bin/activate
```

## 3. Install Requirements

```bash
pip install -r requirements.txt
```

---

# Requirements

```txt
streamlit
openai
numpy
pandas
matplotlib
requests
pillow
arabic-reshaper
python-bidi
```

---

# Run the Application

```bash
streamlit run app.py
```

---

# OpenAI API Key

You must provide your OpenAI API key inside the sidebar.

The key:
- is used only during the session
- is never stored

---

# Technologies Used

- Streamlit
- OpenAI GPT-4o
- GPT-4o Vision
- NumPy
- Pandas
- Matplotlib
- Pillow
- Requests

---

# License

This project is for educational and research purposes.
