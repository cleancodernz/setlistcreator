import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import argparse

# Function to create the setlist document
def create_setlist(csv_file, songs_per_set):
    # Read the CSV file containing songs and lengths
    df = pd.read_csv(csv_file)

    # Initialize a new Word document
    doc = Document()

    # Set document-wide styles for font sizes and alignment
    def add_song_to_document(doc, song, length):
        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = p.add_run(f"{song} - {length}")
        run.font.size = Pt(18)

    # Loop through the songs in sets
    total_songs = len(df)
    num_pages = (total_songs + songs_per_set - 1) // songs_per_set  # Calculate number of pages (sets)
    
    for page in range(num_pages):
        start_idx = page * songs_per_set
        end_idx = min(start_idx + songs_per_set, total_songs)
        set_songs = df.iloc[start_idx:end_idx]

        # Add each song in the set to the document
        total_time = 0
        for _, row in set_songs.iterrows():
            song = row['Song']
            length = row['Length']
            add_song_to_document(doc, song, length)
            
            # Calculate total set length (in minutes)
            minutes, seconds = map(int, length.split(':'))
            total_time += minutes * 60 + seconds

        # Add the total set length at the bottom of each page
        total_minutes = total_time // 60
        total_seconds = total_time % 60
        p = doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = p.add_run(f"Total Set Length: {total_minutes} min {total_seconds} sec")
        run.font.size = Pt(14)

        # Add a page break after each set, except for the last page
        if page < num_pages - 1:
            doc.add_page_break()

    # Save the document
    doc.save('setlist.docx')
    print("Setlist generated and saved as 'setlist.docx'")

# Main function to handle arguments and execution
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generate a setlist from a CSV of songs.')
    parser.add_argument('csv_file', type=str, help='Path to the CSV file containing song names and lengths.')
    parser.add_argument('songs_per_set', type=int, help='Number of songs per set (page).')
    
    args = parser.parse_args()
    
    # Generate the setlist
    create_setlist(args.csv_file, args.songs_per_set)
