# Canva Presentation Screenshot Automation

This Python script automates the process of capturing screenshots from a Canva presentation and converting them into a PowerPoint (.pptx) file.

## 📌 Features

- Automatically opens a Canva presentation in Chrome.
- Enters full-screen "Presentation" mode.
- Captures a specified number of slides using screenshots.
- Saves each slide as a PNG image.
- Compiles all captured slides into a PowerPoint presentation.

## 🧰 Requirements

- Google Chrome
- [ChromeDriver](https://sites.google.com/a/chromium.org/chromedriver/downloads) (make sure it's compatible with your Chrome version)
- Python 3.x
- Python packages:
  - `selenium`
  - `pyautogui`
  - `python-pptx`

## 🔧 Setup

1. Clone this repository or download the script.
2. Install the required Python packages:

   ```bash
   pip install selenium pyautogui python-pptx
   ```

3. Download and extract the ChromeDriver, and update the `CHROMEDRIVER_PATH` variable in the script to point to the driver executable.

4. Set the desired Canva presentation link and the number of slides to capture in the configuration section of the script.

## ⚙️ Usage

Run the script using Python:

```bash
python canva_screenshot_to_ppt.py
```

The script will:
1. Open the Canva link.
2. Enter full-screen presentation mode.
3. Capture the specified number of slides.
4. Save each slide as a PNG in the `screenshots/` directory.
5. Compile the screenshots into a `.pptx` file named `Canva_Presentation_Auto.pptx`.

## 📝 Configuration

Inside the script, you can customize:

```python
CANVA_LINK = 'https://www.canva.com/design/your-presentation-link'
TOTAL_SLIDES = 10
CHROMEDRIVER_PATH = 'path/to/your/chromedriver.exe'
SCREEN_WIDTH = 2560
SCREEN_HEIGHT = 1440
```

Make sure your screen resolution matches `SCREEN_WIDTH` and `SCREEN_HEIGHT` for accurate screenshots.

## ⚠️ Notes

- This script uses `pyautogui`, which relies on the active screen. Do not interrupt the process while it runs.
- It's optimized for use on Windows with a 2560x1440 resolution. Adjust screen dimensions as needed.
- Some Canva presentations may take longer to load; you may need to increase sleep times for reliability.

## 📁 Output

- PNG screenshots saved in the `screenshots/` directory.
- Final PowerPoint saved as `Canva_Presentation_Auto.pptx`.
