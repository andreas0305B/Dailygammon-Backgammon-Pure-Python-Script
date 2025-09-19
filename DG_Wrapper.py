"""
Wrapper Script for DailyGammon Scores

This script loops over a predefined list of leagues and executes the main
script for each league in sequence.

⚠️ Important:
- The main script already supports command-line arguments and auto-closes 
  the Excel workbook when called with '--auto'. 
- For full details about behavior and options, see the docstring 
  in main script.
"""
import subprocess

# List of all leagues that should be processed in one run.
# You can adjust this list as needed (e.g. add/remove leagues).
LIGEN = ["3a", "4d"]

# Name of the main script that performs the actual work for a single league.
MAIN_SCRIPT = "DGscorefetcher.py"

def run_all_leagues():
    """
    Run the main script once for each league defined in LIGEN.
    
    The league name is passed as a command-line argument to the main script.
    
    """
    for liga in LIGEN:
        print(f"\n=== Running league {liga} ===")
        # subprocess.run ensures the script is executed as if started from the shell
        # 'check=True' will raise an error if the script fails
        subprocess.run(["python", MAIN_SCRIPT, liga, "--auto"], check=True) # auto ensures the close of the excel wb in the main script

if __name__ == "__main__":
    run_all_leagues()

