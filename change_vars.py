import re

import vars

unique = []
repeateds = []
for value in vars.replacements.values():
    if value not in unique:
        unique.append(value)
    else:
        repeateds.append(value)

if len(repeateds) != 0:
    print(f"[ERROR] REPEATED VALUES: {repeateds}")
    exit()


def replace_whole_words(text, replacements):
    # Create a regex pattern to match whole words only
    pattern = (
        r"\b("
        + "|".join(re.escape(key.upper()) for key in replacements.keys())
        + r")\b"
    )
    # Replace matched whole words using the dictionary
    return re.sub(
        pattern, lambda match: f"calcs.{replacements[match.group(0).lower()]}", text
    )


# Example usage
text = """

"""

result = replace_whole_words(text, vars.replacements)
print(result)
