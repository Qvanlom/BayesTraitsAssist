import gspread
from oauth2client.service_account import ServiceAccountCredentials
import subprocess
from gspread.exceptions import WorksheetNotFound
import csv
from collections import defaultdict
import itertools

# Ensure names are correct
bayestraits_exe = ".\\BayesTraitsV5"
tree_file = "bayestreefile.trees"
data_file = "output.data"

def run_program(trait1, trait2, sheet2, mode):
    print(f"Running {trait1} vs {trait2}")

    # Collects values
    idx1 = header_row.index(trait1)
    idx2 = header_row.index(trait2)
    col1_vals = [row[idx1] for row in data_rows]
    col2_vals = [row[idx2] for row in data_rows]

    # List of labels
    labels = ["Acamp_Uros",
            "Agath_Agat",
            "Agath_Cocc",
            "Alysi_Alys",
            "Alysi_Colo",
            "Alysi_Dacn",
            "Alysi_Hopl",
            "Aphid_Acan",
            "Aphid_Aphi",
            "Aphid_Ephe",
            "Aphid_Prao",
            "Aphid_Xeno",
            "Brach_ABla",
            "Brach_Dios",
            "Brach_Gryp",
            "Brach_Neal",
            "Brach_Schi",
            "Brach_Vadu",
            "Braco_Vipi",
            "Braco_sp_B",
            "Cardi_Card",
            "Cenoc_Ceno",
            "Charm_Char",
            "Chelo_Chel",
            "Chelo_Phan",
            "Doryc_Hete",
            "Doryc_Noti",
            "Doryc_Onts",
            "Doryc_Rhac",
            "Doryc_Spat",
            "Eupho_Cent",
            "Eupho_Cosm",
            "Eupho_Ecno",
            "Eupho_Elas",
            "Eupho_Leio",
            "Eupho_Mete",
            "Eupho_Micr",
            "Eupho_Zele",
            "Exoth_Cola",
            "Gnamp_Pseu",
            "Helco_Helc",
            "Homol_Homo",
            "Hormi_1Par",
            "Hormi_Aulo",
            "Hormi_Cedr",
            "Hormi_Chre",
            "Hormi_Horm",
            "Ichne_Ichn",
            "Ichne_Olig",
            "Ichne_Prot",
            "Macro_Hyme",
            "Maxfi_Maxf",
            "Mesos_Xeno",
            "Micro_Apan",
            "Micro_Micr",
            "Micro_Phol",
            "Mirac_Mira",
            "Opiin_Bios",
            "Opiin_Fopi",
            "Opiin_Opiu",
            "Orgil_Orgi",
            "Pambo_Pamb",
            "Rhysi_Cant",
            "Rhysi_Rhys",
            "Rhyss_Hist",
            "Rhyss_Onco",
            "Rogad_Alei",
            "Rogad_Clin",
            "Rogad_Stir",
            "Rogad_Trir",
            "Rogad_Yeli",
            "Sigal_Siga",
            "Trach_Trac",
            "BS_UCE_Ich"]

    # Create output.data
    with open("output.data", "w", encoding="utf-8") as f:
        for i in range(len(labels)):
            f.write(f"{labels[i]}\t{col1_vals[i]}\t{col2_vals[i]}\n")

    # Run BayesTraits twice (model 2 and 3)
    results = []
    for model_type in [2, 3]:
        # Run BayesTraits
        run_bayestraits(model_type)

        # Read LML from Stones file
        if (model_type == 2):
            lml2 = read_lml_from_stones("output.data.Stones.txt")
            results.append(float(lml2))
        else:
            lml3 = read_lml_from_stones("output.data.Stones.txt")

            # Read and average q values and root probs
            q_averages, root_averages = read_log_averages("output.data.Log.txt")

            results.extend((float(lml3), q_averages, root_averages))
    
    # Unpack results
    lml2, lml3, q, root = results

    # Calculate Bayes Factor
    bayes_factor = 2 * (lml3 - lml2)

    # Write to next row in Sheet2
    all_values = [f"{trait1}+{trait2}", round(bayes_factor, 5), lml2, lml3] \
                 + [round(q[i], 5) for i in q] \
                 + [round(root[r], 5) for r in root]

    existing = sheet2.col_values(1)
    next_row = len(existing) + 1
    sheet2.update(range_name=f"A{next_row}", values=[all_values])

    # Add color to cell
    cell_color = get_bayes_color(bayes_factor)
    request_body = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet2._properties['sheetId'],
                        "startRowIndex": next_row - 1,   # 0-based index
                        "endRowIndex": next_row,
                        "startColumnIndex": 1,  # Column B = 1
                        "endColumnIndex": 2
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": cell_color
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            }
        ]
    }
    spreadsheet.batch_update(request_body)

    # Progress update
    if(mode == 1):
        print(f"{float(next_row - 1) / len(combinations) * 100:.2f}% Complete")


# Runs BayesTraits, model_type is either 2 or 3
def run_bayestraits(model_type):
    commands = f"""{model_type}
    2
    PriorAll exp 10
    Stones 100 1000
    Info
    run
    """

    process = subprocess.Popen(
        [bayestraits_exe, tree_file, data_file],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        cwd="."
    )

    stdout, stderr = process.communicate(input=commands)

    if "Fatal error" in stderr:
        print("BayesTraits error:\n", stderr)
        raise RuntimeError("BayesTraits failed to read model input.")

# Reads the lml from the stones.txt file
def read_lml_from_stones(file):
    last_number = None

    with open(file, "r", encoding="utf-8") as f:
        lines = f.readlines()
        for line in reversed(lines):
            if "Log marginal likelihood:" in line:
                parts = line.strip().split("\t")
                if len(parts) > 1:
                    last_number = parts[-1]
                break

    # Make sure we got the number before proceeding
    if last_number is None:
        raise ValueError("Could not find the log marginal likelihood in the file.")
    return last_number
    
def read_log_averages(file):
    # Calculate and save values
    q_sums = defaultdict(float)
    q_counts = defaultdict(int)
    q_order = []
    root_keys = ["Root - P(0,0)", "Root - P(0,1)", "Root - P(1,0)", "Root - P(1,1)"]
    root_sums = defaultdict(float)
    root_counts = defaultdict(int)

    with open(file, "r", encoding="utf-8") as f:
        lines = f.readlines()

    # Find the line with headers (starts with "Iteration")
    header_index = None
    for i, line in enumerate(lines):
        if line.strip().startswith("Iteration") and '\t' in line and 'q' in line:
            header_index = i
            break

    if header_index is None:
        raise ValueError("Header line starting with 'Iteration' not found in the log file.")

    # Parse remaining lines with csv.DictReader
    data_lines = lines[header_index:]
    reader = csv.DictReader(data_lines, delimiter='\t')

    # Collect and average q values & root values
    for row in reader:
        for key, value in row.items():
            if key is None or value is None:
                continue
            # Average q-values
            if key.startswith("q"):
                try:
                    val = float(value)
                except ValueError:
                    continue
                if key not in q_order:
                    q_order.append(key)
                q_sums[key] += val
                q_counts[key] += 1
            # Average root values
            elif key.strip() in root_keys:
                try:
                    val = float(value)
                except ValueError:
                    continue
                root_sums[key.strip()] += val
                root_counts[key.strip()] += 1

    q_averages = {key: q_sums[key] / q_counts[key] for key in q_order}
    root_averages = {key: root_sums[key] / root_counts[key] for key in root_keys if key in root_sums}
    return q_averages, root_averages

# Returns the title from row 1 for the given column letter
def get_column_title(sheet, col_letter):
    col_index = ord(col_letter.upper()) - ord('A') + 1  # 1-based index for gspread
    return sheet.cell(1, col_index).value

def get_bayes_color(bf):
    if bf > 10:
        return {"red": 0.0, "green": 0.8, "blue": 0.0}  # Green
    elif bf >= 5:
        return {"red": 1.0, "green": 1.0, "blue": 0.0}  # Yellow
    elif bf > 2:
        return {"red": 1.0, "green": 0.65, "blue": 0.0} # Orange
    else:
        return {"red": 1.0, "green": 0.0, "blue": 0.0}  # Red

# Main starts here-----------------------------------------------------------------------------------
# Authenticate with Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# Ask user for spreadsheet name
spreadsheet_name = input("Enter the name of the Google Sheet (as it appears in Google Drive). Remember to share the sheet with your Google Cloud email first: ").strip()

# Open the Google Sheet
try:
    spreadsheet = client.open(spreadsheet_name)
except gspread.exceptions.SpreadsheetNotFound:
    print(f"Error: Could not find a spreadsheet named '{spreadsheet_name}'.")
    exit(1)

sheet = spreadsheet.sheet1
header_row = sheet.row_values(1)

# Change number below if the traits dont start at column E. A=0,B=1,C=2,D=3,E=4,etc 
# (EX: If the data starts at column C. Line 300: trait_columns = header_row[2:])
trait_columns = header_row[4:]

# Save sheet values
all_rows = sheet.get_all_values()
header_row = all_rows[0]
data_rows = all_rows[1:]

# Try to open "Results" referenced as sheet2, create it if it doesn't exist
try:
    sheet2 = spreadsheet.worksheet("Results")
except WorksheetNotFound:
    sheet2 = spreadsheet.add_worksheet(title="Results", rows="100", cols="10")

sheet2.clear() # Clears sheet
# Adds headers
headers = ["Compared Traits", "Bayes Factor", "Independent LML", "Dependent LML", "q12", "q13", "q21", "q24", "q31", "q34", "q42", "q43", "Root - P(0,0)", "Root - P(0,1)", "Root - P(1,0)", "Root - P(1,1)"]
sheet2.update(range_name="A1:P1", values=[headers])

# Prompt user
print("Select mode:")
print("1 - Run all column combinations")
print("2 - Run specific column pair")
choice = input("Enter 1 or 2: ").strip()

if choice == "1":
    # Loop starts here
    combinations = list(itertools.combinations(trait_columns, 2))
    print("Trait columns:", len(trait_columns))
    print(f"Possible Combinations: {len(combinations)}")
    print(f"Estimated Time: {len(combinations) * 0.6} minutes")
    for trait1, trait2 in combinations:
        run_program(trait1, trait2, sheet2, 1)
    exit(0)

elif choice == "2":
    col1 = input("Enter first column letter (e.g. \"A\"): ").strip().upper()
    col2 = input("Enter second column letter (e.g. \"B\"): ").strip().upper()

    # Fetching column titles from Sheet1
    col1_trait = get_column_title(sheet, col1)
    col2_trait = get_column_title(sheet, col2)

    # Now build .data, run BayesTraits, process results, etc.
    run_program(col1_trait, col2_trait, sheet2, 2)
    print("Task Complete")
    exit(0)

else:
    print("Invalid input. Please run the program again and enter 1 or 2.")
    exit(1)