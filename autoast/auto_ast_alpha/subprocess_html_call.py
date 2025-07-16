import sys
print("[DEBUG] sys.executable:", sys.executable)
print("[DEBUG] sys.path:", sys.path)

from fc_to_html import HTMLGenerator

if __name__ == "__main__":
    HTMLGenerator.run_from_cli()