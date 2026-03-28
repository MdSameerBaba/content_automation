# PDF to Video Pipeline

Convert PDF presentations to MP4 videos with AI narration.

## Prerequisites

- Python 3.10+
- LibreOffice (for PDF to PPTX conversion)
- 8GB+ RAM recommended

## Setup

### 1. Create Virtual Environment

```bash
# Create venv
python -m venv venv

# Activate on Windows
venv\Scripts\activate

# Activate on Mac/Linux
source venv/bin/activate
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Install Additional Dependencies

This project requires some extra packages:

```bash
pip install kokoro imageio-ffmpeg
```

### 4. Environment Variables

Create a `.env` file in the `ppt-pipeline` folder:

```env
# Required - Get from https://platform.openai.com/
OPENAI_API_KEY=your_api_key_here
```

## Running the Application

```bash
cd ppt-pipeline
python app.py
```

Open http://localhost:5000 in your browser.

## Usage

1. **Upload PDF**: Drag & drop a PDF file
2. **Stage 1 - Parse**: Extract slides from PDF
3. **Stage 2 - Structure**: AI groups slides into sections
4. **Stage 3 - Content**: AI adds speaker notes
5. **Stage 4 - Build**: Create themed PPTX
6. **Stage 5 - Images**: Export slides as PNG
7. **Stage 6 - Audio**: Generate voice narration (choose voice first!)
8. **Stage 7 - Video**: Combine into final MP4

## Voice Selection

In Stage 6 (Audio), you can choose from 7 Kokoro AI voices:

| Voice ID | Name |
|----------|------|
| af_heart | Heart (Female) |
| af_sarah | Sarah (Female) |
| am_adam | Adam (Male) |
| am_michael | Michael (Male) |
| bf_emma | Emma (British Female) |
| bm_george | George (British Male) |
| af_nova | Nova (Female) |

Click "Preview" to hear each voice before generating.

## Troubleshooting

### "moov atom not found" error
- Try regenerating the video (Stage 7)

### Kokoro fails with dependency error
```bash
pip install --upgrade click typer
```

### LibreOffice not found
- Install LibreOffice from https://www.libreoffice.org/
- Add to PATH or specify location in code

## Project Structure

```
ppt-pipeline/
├── app.py                 # Flask web app
├── pipeline/
│   ├── stage1_parser.py   # PDF parsing
│   ├── stage2_structurer.py # Slide grouping
│   ├── stage3_content.py   # AI content generation
│   ├── stage4_builder.py   # PPTX creation
│   ├── stage5_images.py    # PNG export
│   ├── stage6_audio.py     # Voice generation
│   └── stage7_video.py    # MP4 creation
├── static/index.html      # Frontend UI
├── theme/                 # PPTX templates
├── uploads/              # Uploaded PDFs
└── checkpoints/          # Processing data
```
