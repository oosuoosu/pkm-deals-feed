import os
from datetime import datetime

# Test saving a simple text file to the specified path
output_dir = r"C:\Users\Philip Work\Documents\Python"
os.makedirs(output_dir, exist_ok=True)

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
file_path = os.path.join(output_dir, f"test_output_{timestamp}.txt")

with open(file_path, "w", encoding="utf-8") as f:
    f.write("This is a test file to confirm write access to the folder.")

file_path
