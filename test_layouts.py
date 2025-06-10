from pptx import Presentation

# Load your PowerPoint template
prs = Presentation("data/3Q/Cat Rock Capital 3Q24 Review Webinar Presentation - Technical Case.pptx")

# Print available slide layouts
print("Available slide layouts in the presentation:")
for i, layout in enumerate(prs.slide_layouts):
    print(f"{i}: {layout.name}")