# Astronomy Olympiad PDF Grader (v1.3)

**Author:** Václav Pavlík  
**Assisted by:** ChatGPT  

Offline grading tool designed for efficient evaluation of PDF submissions for the **Czech Astronomy Olympiad**.

---

## ✨ Overview

This application is optimized for **fast, keyboard-driven grading** of large batches of student PDFs. It emphasizes:

- Minimal mouse-clicking  
- High-speed navigation  
- Persistent grading state  

---

## 🚀 Quick Start

> The tool is designed for **question-by-question** as well as **student-by-student grading**.

### First time

1. Load submissions folder  
2. Define questions  
3. Create a CSV file
4. Load clean answersheet
5. Define an Anchor for each question
6. Start grading
- Select a question and grade all students using `PgUp / PgDown`
- Select a student and grade all questions using `Left / Right`
7. Repeat for the next question or student
8. Save (automatic saves also happen when you select a different student)

### Resume work

1. Load CSV file
2. Optionally load the submissions folder (if it does not load automatically)
3. Start grading

---

## 💻 Controls

| Action | Shortcut |
|------|--------|
| Next / previous ungraded student | `PgDown / PgUp` |
| Sequential navigation | `Shift + PgDown / Shift + PgUp` |
| Change question | `Left / Right` |
| Apply scoring bucket | `0–9` |
| Zoom PDF | `Ctrl + Scroll` |
| Scroll through PDF | `Up / Down` |
| Apply custom score | `Enter` |
| Save | `Ctrl + S` |

---

## 🎨 Visual Guide

**Students:**
- White → Not graded  
- Yellow → Partially graded  
- Green → Fully graded  

**Questions:**
- Green → Completed for all students  

---

## 🧮 Grading Model

Each question uses **buckets (rubric items)**.

Buckets can:
- Add points  
- Set final score  

You can:
- Click buttons  
- Use number shortcuts  

> A custom score overrides bucket selections.

---

## 📍 Anchors

Use **"Set Anchor"** on a clean answer sheet:

1. Select a question
2. Click where the question begins  
3. The app will then automatically scroll student PDFs to this location

---

## 📂 File Structure

```
grades.csv              # Scores, notes, grading status
grades.schema.json      # Configuration, anchors, app state (auto-created)
/some-folder/*.pdf      # Student submissions location
```

### Notes

- `*.schema.json` is **created/updated automatically on save** and loaded from the **same directory as the CSV file**  
- The above structure and names are **recommended**, not required  
- You can freely **navigate and load files from different directories** within the app  
- If paths remain consistent, previously used files may load automatically  

---

## 📦 Portability

The application is **fully portable**:

- You can copy the entire app to another computer and run it immediately  
- Load your existing `grades.csv` and configuration files  
- Load the submissions folder manually  

If the folder structure is preserved, PDFs may be detected automatically.

> No installation or network connection required.

---

## ⚙️ Features

### PDF Handling
- Fast rendering via **PyMuPDF**
- Background preloading
- Render caching for smooth navigation

### Navigation
- Automatic skipping of graded students
- Efficient keyboard workflow

### Grading
- Bucket-based scoring system
- Keyboard shortcuts
- Add / set scoring modes

### Persistence
- Scores, notes, buckets
- Anchors
- Last used paths

---

## 🧠 Tips for Fast Grading

- Prefer keyboard over mouse  
- Grade one question across all students  
- Use anchors to eliminate manual scrolling  

---

## ⚠️ Limitations

- No PDF splitting  
- No cloud synchronization  
- Memory usage increases with caching (internally managed)  

---

## 🔧 Implementation

- Python (Tkinter UI)  
- PyMuPDF (PDF rendering)  
- Pillow (image handling)  
- Threaded prefetching  

---

## 🔮 Future Improvements

- Partial (visible-only) rendering  
- Statistics and export tools  
- Multi-monitor support  

---

## 📌 Notes

- One PDF per student  
- Filename must match student identifier  
- First PDF load may be slower (cached afterward)  
- Fully offline operation  
