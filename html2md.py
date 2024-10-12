import re

# File path (update this path to where your input text file is located)
input_file = r"S:\repos\py\8.md"

# List to store extracted URLs and titles
pdf_data = []

# Regular expression pattern to match the URLs and titles
pattern = r'\[pdf-embedder url="([^"]+)" title="([^"]+)"\]'

# Read the input file and extract data
with open(input_file, "r", encoding="utf-8") as file:
    content = file.read()
    
    # Find all matches of URLs and titles
    matches = re.findall(pattern, content)
    
    # Loop through matches and append to pdf_data list
    for match in matches:
        url, title = match
        pdf_data.append({"title": title, "url": url})

with open("output.md", "w", encoding="utf-8") as f:
    for pdf in pdf_data:
        # Write each link in markdown format
        f.write(f"[{pdf['title']}]({pdf['url']})\n\n")

print("Markdown script generated successfully.")

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