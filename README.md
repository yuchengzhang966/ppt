# PDF to Branded PowerPoint Converter

A Python tool that converts LaTeX-generated PDF presentations into professionally branded PowerPoint presentations. This tool automatically extracts content from PDF slides, applies company branding, and generates polished presentations ready for business use. See Lecture 15 for example.



## Features

- **Automated PDF Processing**: Converts PDF pages to high-quality PNG images
- **Smart Content Detection**: Automatically identifies title pages, table of contents, and content slides
- **Company Branding**: Applies consistent branding using PowerPoint templates
- **Image Optimization**: Automatically trims whitespace and optimizes slide images
- **Flexible Slide Types**: Supports title slides, content slides, subtitle slides, and image slides
- **CSV-Driven Workflow**: Uses CSV files for easy slide management and customization

## Prerequisites

- Python 3.7 or higher
- PowerPoint template file (`.pptx`)

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yuchengzhang966/ppt.git
cd pdf-to-ppt-converter
```

2. Install required dependencies:
```bash
pip install python-pptx pillow PyMuPDF PyPDF2 matplotlib
```

## Project Structure

```
├── main.ipynb              # Main execution notebook
├── new_ppt_util.py         # Updated utility functions
├── ppt_util.py            # Original utility functions
├── template.pptx          # PowerPoint template file
├── output_images/         # Generated PNG images from PDF
└── README.md
```

## Usage

### Basic Workflow

1. **Prepare your files**:
   - Place your PDF file in the project directory
   - Ensure you have a PowerPoint template (`template.pptx`)

2. **Configure paths** in `main.ipynb`:
```python
folder_path = "/path/to/your/project"
pdf_path = os.path.join(folder_path, "your_document.pdf")
pr_path = "/path/to/template.pptx"
title = "Your Presentation Title"
```

3. **Run the conversion**:
   - Open and run `main.ipynb` in Jupyter Notebook
   - The tool will extract slide information and convert PDF pages to images
   - Follow prompts to add custom slide titles if needed

4. **Output**:
   - Generated presentation: `{title}.pptx`
   - Extracted images: `output_images/page_*.png`
   - Slide information: `slide_info.csv`

### Advanced Usage

#### Custom Slide Types

The tool supports different slide types:
- `title page`: Main title slide with title and subtitle
- `table of content`: Table of contents with bullet points
- `subtitle`: Section header slides
- `main`: Content slides with images

#### Manual Slide Customization

You can manually edit the generated CSV file to:
- Change slide titles
- Modify slide types
- Add custom subtitles
- Skip certain pages

#### Batch Processing

For multiple presentations, modify the configuration section and run:
```python
# Process multiple PDFs
pdfs = ["file1.pdf", "file2.pdf", "file3.pdf"]
for pdf in pdfs:
    # Configure and process each PDF
    new_ppt_util.generate_presentation(csv_path, template_path, output_path, folder_path)
```

## Template Configuration

### Setting Up Your PowerPoint Template

Your PowerPoint template (`template.pptx`) must be configured with specific slide layouts. The tool maps different slide types to specific layout indices:

**Required Slide Master Layouts:**
- **Layout 0**: Title slide (for presentation cover page)
- **Layout 1**: Content slide (for table of contents)  
- **Layout 2**: Title-only slide (for section headers/subtitles)
- **Layout 3**: Picture slide (for main content with images)

### How to Configure Template Layouts

1. **Open your template.pptx file in PowerPoint**

2. **Access Slide Master view**:
   - Go to `View` → `Slide Master`
   - You'll see the master slide and layout thumbnails on the left panel

3. **Configure each layout** (ensure they appear in this order):

   **Layout 0 - Title Slide:**
   - Should have: Title placeholder + Subtitle placeholder
   - Used for: Presentation title page
   - Typically includes company logo, branding colors

   **Layout 1 - Content Slide:**  
   - Should have: Title placeholder + Content placeholder (text box)
   - Used for: Table of contents pages
   - Content area should support bullet points/lists

   **Layout 2 - Title Only:**
   - Should have: Title placeholder only
   - Used for: Section dividers and subtitle slides
   - Clean design for chapter/section headers

   **Layout 3 - Picture Slide:**
   - Should have: Title placeholder + Large content area for images
   - Used for: Main content slides with PDF-converted images
   - Image area should be prominent and well-positioned

4. **Save and close** the Slide Master view

5. **Test your template** by running the tool with a sample PDF

### Template Validation

If your slides appear incorrectly, check that:
- Layout indices match the expected order (0, 1, 2, 3)
- Each layout has the required placeholders
- Placeholder positions work well with your content

### Custom Layout Mapping

If you need different layout indices, modify the code in `new_ppt_util.py`:
```python
# In add_title_slide()
fengmian = prs.slide_layouts[0]  # Change index for title slides

# In add_content_slide()  
mulu = prs.slide_layouts[1]      # Change index for content slides

# In add_title_only_slide()
biaoti = prs.slide_layouts[2]    # Change index for title-only slides

# In add_image_slide()
tu = prs.slide_layouts[3]        # Change index for picture slides
```

## Key Functions

### `extract_slide_info(pdf_path, csv_path, title)`
Extracts text from PDF pages and creates a CSV file with slide information.

### `pdf_to_png_and_trim(pdf_path, output_folder)`
Converts PDF pages to PNG images and automatically trims whitespace.

### `generate_presentation(csv_path, template_path, output_path, folder_path)`
Creates the final PowerPoint presentation based on CSV configuration.

### `trim_space(image_path, output_path)`
Removes whitespace from images and optimizes for presentation use.

## Customization

### Image Positioning
Modify the positioning parameters in `add_image_slide()`:
```python
left = (slide_width - new_width) / 2 + Pt(100)  # Horizontal offset
top = (slide_height - new_height) / 2 + Pt(90)   # Vertical offset
```

### Trimming Behavior
Adjust whitespace trimming in `trim_space()`:
```python
box = (5, 100, width - 5, height - 66)  # Crop box coordinates
white_threshold = 100  # RGB threshold for white detection
```

## Troubleshooting

### Common Issues

1. **Missing dependencies**: Install all required packages using pip
2. **Template not found**: Ensure your PowerPoint template path is correct
3. **PDF conversion errors**: Check if your PDF is not password-protected
4. **Image quality issues**: Adjust the DPI parameter in `pdf_to_png_and_trim()`

### Error Handling

The tool includes basic error handling for:
- Invalid file paths
- Missing images
- PDF processing errors
- PowerPoint template issues

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Built with [python-pptx](https://python-pptx.readthedocs.io/) for PowerPoint manipulation
- Uses [PyMuPDF](https://pymupdf.readthedocs.io/) for PDF processing
- Image processing powered by [Pillow](https://pillow.readthedocs.io/)

## Support

If you encounter any issues or have questions, please:
1. Check the troubleshooting section above
2. Search existing issues on GitHub
3. Create a new issue with detailed information about your problem

---


**Note**: This tool is designed for converting academic/research PDFs to business presentations. Results may vary depending on the PDF structure and content layout.


