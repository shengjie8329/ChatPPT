from pptx import Presentation

# Function to load a presentation template
def load_template(template_path: str) -> Presentation:
    # Create a Presentation object from the template path
    prs = Presentation(template_path)
    # Return the Presentation object
    return prs

# Function to get a mapping of slide layouts in the presentation
def get_layout_mapping(prs: Presentation) -> dict:
    # Create an empty dictionary to store the layout mapping
    layout_mapping = {}
    # Iterate through the slide layouts in the presentation
    for idx, layout in enumerate(prs.slide_layouts):
        # Add the layout name and index to the dictionary
        layout_mapping[layout.name] = idx
    # Return the layout mapping dictionary
    return layout_mapping

# Function to print the slide layouts in the presentation
def print_layouts(prs: Presentation):
    # Iterate through the slide layouts in the presentation
    for idx, layout in enumerate(prs.slide_layouts):
        # Print the layout index and name
        print(f"Layout {idx}: {layout.name}")