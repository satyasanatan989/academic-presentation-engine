import subprocess
from pathlib import Path

pptx = Path(r"C:\Users\prabh\Documents\antigravity\PPT\academic-presentation-engine\outputs\PRJ-PHYTOCHEMICALS-001\phytochemicals_biological_functionality.pptx")
out_dir = Path(r"C:\Users\prabh\Documents\antigravity\PPT\academic-presentation-engine\outputs\PRJ-PHYTOCHEMICALS-001\slide_previews")
out_dir.mkdir(parents=True, exist_ok=True)

ps = f"""
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('{str(pptx)}', 2, 0, 0)
$pres.Export('{str(out_dir)}', 'PNG', 1920, 1080)
$pres.Close()
$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Output "Exported"
"""

subprocess.run(["powershell.exe", "-NoProfile", "-Command", ps], check=True)
print("Preview export complete.")
