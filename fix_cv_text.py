import sys
import os

filepath = '/mnt/c/Users/omarc/Desktop/APOLLO/email_outreach_app.py'
with open(filepath, 'r') as f:
    lines = f.readlines()

with open(filepath, 'w') as f:
    for line in lines:
        if line.strip() == '"""' and 'cv_text' in line:
            # Check if it is the rogue line
             pass 
        elif line.startswith('"""') and '{cv_text}' in line:
             # Logic to detect the specific misplaced string
             pass
        # simpler approach: replace the specific pattern
        f.write(line)

print("Actually, I will use sed for a more surgical fix if I can identify the pattern precisely.")
