# GPS to KM - Travel History Report Generator & Vehicle Logbook

A Python desktop application that processes GPS Travel History Excel files exported from vehicle tracking systems. It generates per-vehicle, per-day distance (KM) summary reports with persistent odometer logbook management.

---

## 🚀 Features

- **Automated Travel Processing**: Parses raw GPS travel logs into structured, daily mileage summaries per vehicle.
- **Persistent Odometer Logbook**: Automatically tracks cumulative odometer readings across multiple files and updates with JSON-based persistence.
- **Clean Excel Reports**: Formats generated Excel reports with stylized headers, totals, date ranges, and borders using `openpyxl`.
- **Desktop GUI**: User-friendly Tkinter-based interface with dark/light visual cues, progress tracking, and file status feedback.
- **Drag & Drop Support**: Seamless file input via `tkinterdnd2` (falls back gracefully to standard file dialog if not installed).

---

## 🛠️ Requirements

- **Python 3.8+**
- Dependencies:
  ```bash
  pip install -r requirements.txt
  ```

---

## 📦 Installation & Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/gourav-saharan/GPS-to-KM.git
   cd GPS-to-KM
   ```

2. **Install requirements**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the Application**:
   - Double-click `run_app.bat` (Windows)
   - Or run via terminal:
     ```bash
     python gps_to_km.py
     ```

---

## 📂 File Structure

```
GPS-to-KM/
├── gps_to_km.py              # Main application script & GUI
├── vehicle_odometers.json    # Persistent vehicle odometer database
├── run_app.bat               # Windows batch launcher
├── requirements.txt          # Python dependencies
├── .gitignore                # Git ignore rules
└── README.md                 # Project documentation
```

---

## 📄 License

This project is open source and available under the [MIT License](LICENSE).
