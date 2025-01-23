import os
from PyPDF2 import PdfReader, PdfWriter
import xlsxwriter

def main():
    # Step 1: Get user input for PDF file name (without extension)
    #pdf_name = input("Enter the PDF file name without extension: ")
    pdf_name = './Adjudicación Dabigatrán_updated'
    pdf_path = f".\\{pdf_name}.pdf"
    
    # Step 2: Verify the PDF file exists
    if not os.path.isfile(pdf_path):
        print(f"The file '{pdf_path}' does not exist.")
        return

    # Step 3: Load PDF and get bookmarks
    pdf = PdfReader(pdf_path)
    bookmarks = pdf.outline

    # Step 4: Get the user-provided bookmark names
    #user_bookmark_names = input("Enter the bookmark names separated by '|': ").split('|')
    # Main script
    md_file = './Bookmarks.md'  # Path to your .md file
    user_bookmark_names = load_bookmarks(md_file)

    if user_bookmark_names:
        print("Bookmarks loaded successfully:")
        for i, bookmark in enumerate(user_bookmark_names, start=1):
            print(f"{i}. {bookmark}")
    else:
        print("No bookmarks found or the file is empty.")    
    user_bookmark_names = [name.strip() for name in user_bookmark_names if name.strip()]  # Clean input

    # Step 5: Validate the number of bookmarks
    if len(bookmarks) != len(user_bookmark_names):
        print(f"Mismatch: The PDF has {len(bookmarks)} bookmarks, but you provided {len(user_bookmark_names)} names.")
        return
    
    # Ensure the output folder exists
    output_folder = ".\\Output"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Step 6: Call the split function with the new filenames
    split_pdf_by_bookmarks(pdf_path, output_folder, user_bookmark_names)
    print("PDF split by bookmarks and saved with the specified names.")


def load_bookmarks(file_path):
    """
    Load bookmarks from a file where they are stored as a single line separated by '|'.

    Args:
        file_path (str): Path to the `.md` file.

    Returns:
        list: A list of bookmark names.
    """
    if not os.path.exists(file_path):
        print(f"Error: The file '{file_path}' does not exist.")
        return []

    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read().strip()  # Read the content and strip extra spaces/newlines

    # Split the content by '|'
    return content.split('|')




# Updated split function to accept custom names
def split_pdf_by_bookmarks(path_to_pdf, output_folder, bookmark_names):
    pdf = PdfReader(path_to_pdf)
    bookmarks = pdf.outline
    prev_bookmark_title = None
    prev_bookmark_page_index = None
    root_folder = ".\\"
    # Prepare Excel output
    workbook = xlsxwriter.Workbook(os.path.join(root_folder, 'Bookmarks_exported.xlsx'))
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'PDF File Name')
    row = 1  # Start writing PDF names in Excel from row 1

    # Iterate through bookmarks
    for i, bookmark in enumerate(bookmarks):
        if hasattr(bookmark, 'title'):
            page_index = find_page_index(pdf, bookmark)
            title = sanitize_filename(bookmark_names[i])  # Use user-defined name
            
            if prev_bookmark_title is not None:
                pdf_writer = PdfWriter()
                for j in range(prev_bookmark_page_index, page_index):
                    pdf_writer.add_page(pdf.pages[j])
                output_path = os.path.join(output_folder, f"{prev_bookmark_title}.pdf")
                with open(output_path, "wb") as f_out:
                    pdf_writer.write(f_out)
                
                # Save file name to Excel
                worksheet.write(row, 0, f"{prev_bookmark_title}.pdf")
                row += 1

            prev_bookmark_title = title
            prev_bookmark_page_index = page_index

    # Write the last section
    if prev_bookmark_title is not None:
        pdf_writer = PdfWriter()
        for j in range(prev_bookmark_page_index, len(pdf.pages)):
            pdf_writer.add_page(pdf.pages[j])
        output_path = os.path.join(output_folder, f"{prev_bookmark_title}.pdf")
        with open(output_path, "wb") as f_out:
            pdf_writer.write(f_out)

        # Save last PDF name to Excel
        worksheet.write(row, 0, f"{prev_bookmark_title}.pdf")

    workbook.close()
    print("All bookmarks have been split and saved.")

def sanitize_filename(name):
    # Remove any characters that could cause issues in filenames
    return "".join(c for c in name if c.isalnum() or c in (" ", "-", "_")).rstrip()

def find_page_index(pdf, bookmark):
    # Retrieve page index from bookmark reference
    try:
        return pdf.get_page_number(bookmark.page) if hasattr(bookmark, 'page') else 0
    except Exception as e:
        print(f"Error retrieving page index for bookmark: {bookmark}. Error: {e}")
        return 0

# Run the main function
if __name__ == "__main__":
    main()
