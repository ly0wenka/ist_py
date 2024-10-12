import re

# Open the file and read its content
with open('output.md', 'r', encoding='utf-8') as file:
    content = file.read()

# Replace underscores with spaces only inside square brackets
updated_content = re.sub(r'\[(.*?)\]', lambda x: x.group(0).replace('_', ' '), content)

# Print or save the updated content
print(updated_content)

# Optional: Save the updated content to a new file
with open('updated_file.md', 'w', encoding='utf-8') as updated_file:
    updated_file.write(updated_content)