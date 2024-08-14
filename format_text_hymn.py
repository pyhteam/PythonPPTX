import json
import re


def format_text(data):
    formatted_data = ""
    lines = data.split("\n")
    current_section = ""
    
    for line in lines:
        stripped_line = line.strip()
        
        if stripped_line == "":
            if current_section:
                formatted_data += current_section.strip() + "\n\n"
                current_section = ""
        elif stripped_line.isdigit() or (stripped_line and stripped_line[0].isdigit() and stripped_line[1] == '.'):
            if current_section:
                formatted_data += current_section.strip() + "\n\n" 
                current_section = stripped_line + "\n"
            else:
                current_section = stripped_line + "\n"
        else:
            current_section += stripped_line + " "

    if current_section:
        formatted_data += current_section.strip() + "\n"

    return formatted_data.strip()

def format_text_to_json(data):
    songs = []
    current_song = None
    song_id = 1
    verse_id = 1
    
    lines = data.split('\n')
    for line in lines:
        stripped_line = line.strip()
        
        # Bỏ qua các dòng trống
        if not stripped_line:
            continue
        
        # Kiểm tra xem dòng có định dạng số không
        match = re.match(r'(\d+)\.\s*(.*)', stripped_line)
        if match:
            if current_song:
                songs.append(current_song)
            song_label = int(match.group(1))
            song_name = match.group(2).strip()
            current_song = {
                "id": song_id,
                "label": song_label,
                "name": song_name,
                "verses": []
            }
            song_id += 1
            verse_id = 1
        else:
            if current_song:
                current_song["verses"].append({
                    "song_id": current_song["id"],
                    "id": verse_id,
                    "content": stripped_line
                })
                verse_id += 1
    
    if current_song:
        songs.append(current_song)
    
    return json.dumps(songs, indent=4, ensure_ascii=False)

# run main
if __name__ == "__main__":
    filePath ="data\\Cov Ntseeg Yexus Phoo Nkauj - Hmoob Ntsuab.txt"
    fileName =""
    try:
        with open(filePath, 'r', encoding='utf-8') as file:
            fileName = file.name
            data = file.read()
    except UnicodeDecodeError:
        with open(filePath, 'r', encoding='latin1') as file:
            data = file.read()

    formatted_text = format_text(data)
    # save formatted text to file
    with open(fileName+"_format.txt", "w", encoding='utf-8') as file:
        file.write(formatted_text)
        print("Formatted text saved to successfully.")
    
    # save formatted text to JSON
    formatted_json = format_text_to_json(formatted_text)
    with open(fileName+"_format.json", "w", encoding='utf-8') as file:
        file.write(formatted_json)
        print("Formatted json saved to successfully.")
    print("All done!")

