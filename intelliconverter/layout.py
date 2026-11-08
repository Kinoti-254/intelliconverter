"""Page-layout candidates and the layout score."""

# All layout candidates with their printable widths in inches
LAYOUTS = [
    ("A4 Portrait",      7.7),
    ("A4 Landscape",    10.7),
    ("A3 Portrait",     11.7),
    ("Legal Landscape", 13.2),
    ("A3 Landscape",    16.5),
]

def score_layout(total_width, page_width):
    """
    Score a layout option. Returns -1 if content doesn't fit.
    Higher score = better layout.
    
    Scoring logic:
    - Content must fit within page width (otherwise score = -1)
    - Best score when content fills 75-95% of the page
    - Penalise layouts where content is too small (lots of empty space)
    """
    if total_width > page_width:
        return -1  # Doesn't fit — disqualified

    scale = total_width / page_width

    if scale >= 0.75:
        # Sweet spot: content fills most of the page
        return round(50 + (scale * 50), 1)
    else:
        # Too much empty space — lower score
        return round(scale * 60, 1)
