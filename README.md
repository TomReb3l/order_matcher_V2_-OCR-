# Order Matcher

Desktop εφαρμογή για σύγκριση στοιχείων μεταξύ **PDF διαταγής** και **Excel δύναμης υπηρεσίας**.

Υποστηρίζει δύο εκδόσεις:

- **OrderMatcher-Lite**: χωρίς OCR
- **OrderMatcher-OCR**: με ενσωματωμένο Tesseract για OCR σε σαρωμένα PDF

---

## Τι κάνει

Η εφαρμογή:

- φορτώνει PDF διαταγής
- φορτώνει Excel δύναμης
- εντοπίζει κοινά πρόσωπα
- εμφανίζει:
  - κοινές εγγραφές
  - εγγραφές μόνο στο PDF
  - εγγραφές μόνο στο Excel
  - καταγραφές / logs
- κάνει εξαγωγή σε Word
- κάνει εξαγωγή και σε Excel

---

## Εκδόσεις

### Lite

Η έκδοση **Lite** προορίζεται για PDF που περιέχουν ήδη αναγνώσιμο κείμενο.

Δεν χρησιμοποιεί OCR.

### OCR

Η έκδοση **OCR** προορίζεται για σαρωμένα ή image-based PDF.

Χρησιμοποιεί bundled **Tesseract** από τον φάκελο:

`third_party/tesseract/`

Η OCR υλοποίηση καλεί native το `tesseract.exe` και διαβάζει TSV output.

---

## Δομή repo

```text
order_matcher/
├─ .gitignore
├─ LICENSE
├─ README.md
├─ app.py
├─ core.py
├─ build_config.py
├─ build_lite.bat
├─ build_ocr.bat
├─ make_portable.bat
├─ installer_lite.iss
├─ installer_ocr.iss
├─ requirements-lite.txt
├─ requirements-ocr.txt
└─ third_party/
   └─ tesseract/
      ├─ tesseract.exe
      ├─ *.dll
      └─ tessdata/
         ├─ ell.traineddata
         ├─ eng.traineddata
         └─ osd.traineddata
```

---

## Τι δεν ανεβαίνει στο repo

Τα παρακάτω είναι build / local artifacts και καλό είναι να μένουν εκτός Git:

```text
.venv/
.venv_lite/
.venv_ocr/
build/
dist/
portable/
installer_output/
__pycache__/
*.pyc
*.pyo
*.pyd
*.spec
```

---

## Απαιτήσεις

- Windows
- Python 3.11 ή 3.12 προτεινόμενο
- Inno Setup 6 μόνο αν θέλεις installer `.exe`

---

## Build - Lite

Τρέξε:

```bat
build_lite.bat
```

Παράγει:

`dist/OrderMatcher-Lite/`

Το script καθαρίζει μόνο το `build\OrderMatcher-Lite` και το `dist\OrderMatcher-Lite`, χωρίς να σβήνει τυχόν υπάρχον OCR build.

---

## Build - OCR

Τρέξε:

```bat
build_ocr.bat
```

Παράγει:

`dist/OrderMatcher-OCR/`

Το script καθαρίζει μόνο το `build\OrderMatcher-OCR` και το `dist\OrderMatcher-OCR`, χωρίς να σβήνει τυχόν υπάρχον Lite build.

---

## Προτεινόμενη σειρά build

Για να έχεις και τις δύο εκδόσεις διαθέσιμες ταυτόχρονα:

1. Τρέχεις `build_lite.bat`
2. Τρέχεις `build_ocr.bat`

ή και ανάποδα. Τα δύο scripts πλέον δεν αλληλοσβήνονται στο `dist/`.

---

## Portable πακέτα

Αφού ολοκληρωθούν τα builds, τρέξε:

```bat
make_portable.bat
```

Θα δημιουργήσει στον φάκελο `portable/`:

- `OrderMatcher-Lite-Portable.zip`
- `OrderMatcher-OCR-Portable.zip`

για όσα builds υπάρχουν μέσα στο `dist/`.

Τα ZIP περιέχουν τον αντίστοιχο φάκελο εφαρμογής ως root folder, ώστε ο χρήστης να κάνει extract πιο καθαρά.

---

## Installer

Αν έχεις εγκατεστημένο **Inno Setup 6**, μπορείς να χτίσεις installers με:

- `installer_lite.iss`
- `installer_ocr.iss`

Θα χρησιμοποιήσουν αντίστοιχα τους φακέλους:

- `dist/OrderMatcher-Lite/`
- `dist/OrderMatcher-OCR/`

και θα παράγουν installer `.exe` στον φάκελο `installer_output/`.

---

## GitHub Releases

Αν θέλεις πλήρες release, τα πιο χρήσιμα artifacts είναι:

- `OrderMatcher-Lite`
- `OrderMatcher-OCR`
- `OrderMatcher-Lite-Portable.zip`
- `OrderMatcher-OCR-Portable.zip`

και προαιρετικά:

- `OrderMatcher-Lite-Setup.exe`
- `OrderMatcher-OCR-Setup.exe`

---

## Σημειώσεις

- Το `build_config.py` πρέπει να μένει default σε **Lite** mode μέσα στο repo.
- Δεν χρειάζεται `.spec` αρχείο στο τρέχον setup.
- Η OCR έκδοση χρησιμοποιεί bundled Tesseract από το repo.
- Η τελική OCR ροή παράγει TSV μέσω native κλήσης του `tesseract.exe` με `tessedit_create_tsv=1`.

---

## Άδεια

Βλέπε το αρχείο `LICENSE`.
