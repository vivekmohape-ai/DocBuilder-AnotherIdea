# DocBuilder — Real Estate Image Portfolio Generator

A Streamlit app that converts a ZIP of project images into a formatted Word document, using a pasted sequence list to control order, grouping, and captions.

---

## Features

- Upload a ZIP with any subfolder structure (project-wise)
- Paste a simple text sequence to define image order, sections & captions
- Auto-detects orientation: landscape images pair side-by-side, portrait images display centred
- Generates a clean, professionally formatted `.docx` in one click
- Handles 150–200+ images efficiently

---

## How to Deploy on Streamlit Cloud

1. Push this folder to a GitHub repository (public or private)
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app**
4. Select your repo, branch, and set **Main file path** to `app.py`
5. Click **Deploy** — live in ~60 seconds

No server setup, no cost (free tier supports this easily).

---

## Flow Text Format

```
[Project Section Name]
filename.jpg | Caption for this image
filename2.png | Another caption

[Next Project]
image_a.jpg | Caption A
image_b.jpg | Caption B
image_c.jpg | Caption C
```

- Section headers use `[Square Brackets]`
- Each image line: `filename | caption`  (separator can be `|`, `-`, or `,`)
- Filenames are matched case-insensitively
- Partial filename matching is supported (e.g. `ooh` will match `empress_ooh_final.jpg`)

---

## Layout Logic

| Image orientation | Layout |
|---|---|
| Landscape + next image also landscape | Side-by-side (2 per row) |
| Portrait | Centred single column |
| Landscape with no landscape neighbour | Full width |

---

## ZIP Structure

Any structure works:
```
images.zip
├── Empress_Hill/
│   ├── digital_screen_1.jpg
│   ├── ooh_poster.jpg
│   └── standee.jpg
├── Safety_Week/
│   ├── badge.png
│   └── banner.jpg
└── Enquiry_Forms/
    └── enquiry_form.jpg
```

---

## Local Run

```bash
pip install -r requirements.txt
streamlit run app.py
```
