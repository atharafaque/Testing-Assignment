import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {
    "packages": [
        "os", 
        "selenium",
        "selenium.webdriver",
        "selenium.webdriver.common",
        "selenium.webdriver.support",
        "pandas",
        "webdriver_manager",
        "logging",
        "traceback",
        "random",
        "time",
        "pathlib"
    ],
    "includes": [
        "selenium.webdriver.chrome.service",
        "selenium.webdriver.common.by",
        "selenium.webdriver.common.keys",
        "selenium.webdriver.support.ui",
        "selenium.webdriver.support.expected_conditions",
        "webdriver_manager.chrome"
    ],
    "excludes": [],
    "include_files": [
        "test_data.xlsx",
        "my_resume.pdf",
        "gmail_automation.log",
        ("screenshots", "screenshots")
    ]
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="GmailAutomation",
    version="1.0",
    description="Gmail Automation Tool",
    options={"build_exe": build_exe_options},
    executables=[Executable("gmail_automation_poc.py", base=base)]
)