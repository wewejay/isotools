from isoarchive import *


outfile = "exes.json"
outputsep = "--------------------"
mountediso = ""

print("Loading JSON...")
try:
    with open(outfile) as file:
        ordict = json.load(file)
except:
    print("File not found or empty.")
    input("Press Enter to terminate Program...")
    sys.exit()
orlist = list(ordict.items())
print("JSON Loaded.")
print(outputsep)
results = []
while True:
    termsinput = input("Input(? for help): ")
    if not termsinput.strip():
        print("Empty input not allowed.")
        continue
    if termsinput == '?q':
        break
    if termsinput == '?':
        print("\ntype search terms to search (eg. \"net .exe\")"
              "\n? to show this"
              "\n?q to quit"
              "\n?<x> : get the x-th file/directory in search result\n")
        continue
    if termsinput.startswith('?'):
        wantedindex = 0
        try:
            wantedindex = int(termsinput.lstrip('?'))
        except:
            print("Not a valid input.")
            continue
        if not results:
            print("Search before mounting!")
            continue
        if wantedindex >= len(results):
            print("File index out of range.")
            continue
        isopath, filepath = getisoandfile(results, orlist, wantedindex)
        if not mountandget(isopath, filepath):
            print("Error")
            continue
        i_temp = input("\nType anything to dismount, s to continue without dismounting: ")
        if i_temp == 's':
            print("Skipped dismounting.")
            print(outputsep)
            continue
        print("\nDismounting...")
        if not dismount(isopath):
            print("Error: Please dismount manually!")
            continue
        print("Dismounted.")
    else:
        searchterms = termsinput.split()
        results = search(orlist, searchterms)
        print()
        printresult(results, orlist, 2)
    print(outputsep)
