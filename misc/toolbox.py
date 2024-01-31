def transform_list(input_list):
    # Séparer les listes non vides des listes vides
    chunks = []
    chunk = []
    for lst in input_list:
        if lst:
            chunk.append(lst)
        else:
            if chunk:
                chunks.append(chunk)
                chunk = []
    if chunk:
        chunks.append(chunk)

    # Fusionner les listes dans chaque chunk
    merged = []
    for chunk in chunks:
        merged_chunk = []
        for lst in chunk:
            merged_chunk.extend(lst)
            merged_chunk.append('')
        merged.append(merged_chunk[:-1])  # Supprimer le dernier espace vide ajouté

    # Ajouter des espaces vides pour que les listes aient la même longueur
    max_length = max(len(lst) for lst in merged)
    for lst in merged:
        while len(lst) < max_length:
            lst.append('')

    # Insérer une liste vide entre les deux parties
    mid = len(merged) // 2
    merged.insert(mid, [])

    # Fusionner les deux premières listes
    merged[0] = merged[0] + [''] + merged[1]
    del merged[1]

    return merged

test = [
    ["a", "a", "a"],
    ["a", "a", "a"],
    ["a", "a", "a"],
    [],
    ["b", "b", "b"],
    [],
    ["c", "c", "c"],
    ["c", "c", "c"],
    ["c", "c", "c"],
    ["c", "c", "c"],
    [],
    ["d", "d", "d"],
    [],
    ["e", "e", "e"],
    ["e", "e", "e"]
]

from pprint import pprint
pprint(transform_list(test))