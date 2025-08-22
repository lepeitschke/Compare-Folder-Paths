import os
import pandas as pd


def get_all_paths(folder):
    """Recursively collect all file paths in a folder with absolute paths."""
    paths = []
    for root, _, files in os.walk(folder):
        for file in files:
            abs_path = os.path.abspath(os.path.join(root, file))
            paths.append(abs_path)
    return set(paths)


def strip_paths(path, string_to_strip):
    stripped = set()
    for i in path:
        stripped.add(i.replace(string_to_strip, ""))
    return stripped


def compare_folders(folder1, folder2, output_file="comparison.xlsx"):
    # Collect all absolute paths
    paths1 = get_all_paths(folder1)
    paths2 = get_all_paths(folder2)

    stripped1 = strip_paths(paths1, "F:\\Root\\Restore-E\\")
    stripped2 = strip_paths(
        paths2, "\\\\NAS-STUDIO\\Recovered Files 12TB HDD\\Restore-E\\"
    )

    # Common files (exact same full path strings)
    common = sorted(stripped1.intersection(stripped2))

    # Files missing in either folder
    only_in_folder1 = sorted(stripped1 - stripped2)
    only_in_folder2 = sorted(stripped2 - stripped1)

    # Prepare DataFrames
    df_common = pd.DataFrame(common, columns=["Common Paths"])

    df_missing = pd.DataFrame(
        {
            f"Only in {os.path.abspath(folder1)}": pd.Series(only_in_folder1),
            f"Only in {os.path.abspath(folder2)}": pd.Series(only_in_folder2),
        }
    )

    # Write to Excel with two worksheets
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_common.to_excel(writer, sheet_name="Common", index=False)
        df_missing.to_excel(writer, sheet_name="Missing", index=False)

    print(f"Comparison saved to {output_file}")


# Example usage
folder_a = r"F:\Root\Restore-E"
folder_b = r"\\NAS-STUDIO\Recovered Files 12TB HDD\Restore-E"
compare_folders(folder_a, folder_b, "folder_comparison.xlsx")
print("DONE!")
