def det_signif_headers(headers):
    """
    Wählt die wichtigsten Elemente aus einer Liste von Überschriften aus.

    Args:
    headers (list): Eine Liste von Strings, die die Überschriften darstellen.

    Returns:
    list: Eine kürzere Liste mit den ausgewählten wichtigen Überschriften.

    Die Funktion wählt alle Strings aus, die "_total" enthalten, sowie den ersten
    String, der "index_" im Namen hat.
    """
    significant_headers = [header for header in headers if "_total" in header]
    
    index_header = next((header for header in headers if "index_" in header), None)
    if index_header:
        significant_headers.append(index_header)
    
    return significant_headers
