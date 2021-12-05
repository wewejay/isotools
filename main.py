import os
import subprocess
import multiprocessing
import time
import datetime
import json
import sys
import winsound


def psrun(cmd):
    completed = subprocess.run(["powershell", "-Command", cmd], capture_output=True)
    return completed


def mount(path):
    info = psrun("$result = Mount-DiskImage \"{}\" -PassThru ; $result | Get-Volume".format(path))
    if info.stderr != b'':
        print(info.stderr.decode("utf-8"))
        return False
    drive = list(str(info.stdout).split()[14])[-1]
    return drive + ":\\"


def dismount(path):
    info = psrun("Dismount-DiskImage \"{}\"".format(path))
    if info.stderr != b'':
        print(info.stderr.decode("utf-8"))
        return False
    return True


def loadjson(outfile):
    with open(outfile) as file:
        ordict = json.load(file)
    return list(ordict.items())


def getisoandfile(results, orlist, index):
    isoindex = results[index][0]
    fileindex = results[index][1]
    isopath = orlist[isoindex][0]
    filepath = orlist[isoindex][1][fileindex]
    return isopath, filepath


def search(orlist, searchterms):
    resultindexlist = []
    # resultlist = []
    for i in range(len(orlist)):
        # isoname = orlist[i][0]
        filepaths = orlist[i][1]
        for j in range(len(filepaths)):
            filepath = filepaths[j]
            filename = extractpath(filepath, 1)
            filteredfilename = ''.join(filename.lower().split())
            found = True
            for term in searchterms:
                if term.lower() not in filteredfilename:
                    found = False
                    break
            if found:
                # print(filename)
                resultindexlist.append((i, j))
                # resultlist.append((isoname, filepath, filename))

    return resultindexlist


def extractpath(fullpath: str, pathlength):
    if pathlength == 0:
        return fullpath
    pathlist = fullpath.split("\\")
    if len(pathlist) <= pathlength:
        return fullpath
    return "\\".join(pathlist[-pathlength:])


def printresult(result, orlist, pathlength):
    if not result:
        print("No Results.")
    else:
        for i in range(len(result)):
            isopath, filepath = getisoandfile(result, orlist, i)
            isoname = extractpath(isopath, 1)
            filename = extractpath(filepath, pathlength)
            col1 = 5
            col2 = 100
            col3 = 20
            plholder = ' '
            indexstr = str(i).ljust(col1, plholder)
            isonamestr = isoname.upper().ljust(col3, plholder)
            filenamestr = filename.ljust(col2, plholder)
            print("{} | {} | {}".format(indexstr, filenamestr, isonamestr))
        print("\n {} Results.".format(len(result)))


def mountandget(isopath, filepath):
    print("Mounting...")
    drivepath = mount(isopath)
    if not drivepath:
        return False
    print("\"{}\" mounted on \"{}\"".format(isopath, drivepath))
    subprocess.Popen("explorer /select,\"{}\"".format(drivepath + filepath))
    print("Redirected to \"{}\"".format(drivepath + filepath))
    return True


def findexe(cdroot):
    allfilelist = []
    for root, dirs, files in os.walk(cdroot):
        for dir_ in dirs:
            allfilelist.append(os.path.relpath(os.path.join(root, dir_), cdroot))
        for file in files:
            pass
            allfilelist.append(os.path.relpath(os.path.join(root, file), cdroot))
    return allfilelist


def process(isonum, isopathlist):
    error = False
    alle = []
    try:
        path = isopathlist[isonum]
        drive = mount(path)
        if not drive:
            raise Exception
        alle = findexe(drive)
        if not dismount(path):
            raise Exception
    except:
        error = True
    return error, alle


def feedback(name, i, isocount, starttime, error):
    errorstr = ""
    plholder = " "
    col1 = 50
    col2 = 12
    col3 = 20
    namestr = name.upper().ljust(col1, plholder)
    progressstr = (str(round(i / isocount * 100, 1)) + '%').ljust(col2, plholder)
    passedtimestr = str(datetime.timedelta(seconds=round(time.time() - starttime))).ljust(col3, plholder)
    if error:
        errorstr = "Error"
    print("{}|{}|{}{}".format(namestr, progressstr, passedtimestr, errorstr))


def parprocess(isocount, isopathlist, isofilelist, outfile, starttime):
    ereturn = []
    finished = 0
    processlist = []
    outdict = {}
    with multiprocessing.Pool() as pool:
        for i in range(isocount):
            processlist.append(pool.apply_async(process, (i, isopathlist)))
        for workprocessnum in range(len(processlist)):
            er, alle = processlist[workprocessnum].get()
            name = isopathlist[workprocessnum]
            outdict.update({name: alle})
            if er:
                ereturn.append(name)
            finished += 1
            feedback(isofilelist[workprocessnum], finished, isocount, starttime, er)
    data = {}
    if os.path.exists(outfile):
        if os.stat(outfile).st_size != 0:
            with open(outfile) as f:
                data = json.load(f)
    data.update(outdict)
    with open(outfile, 'w') as f:
        json.dump(data, f, indent=4)
    return ereturn


def indexer():
    isocount = 0
    isopathlist = []
    isofilelist = []
    isoextensions = ".iso"  # more in tuple
    outfile = "exes.json"
    outputsep = "--------------------"

    isoroot = input("Root for finding ISOs: ")

    for root, dirs, files in os.walk(isoroot):
        for file in files:
            if file.lower().endswith(isoextensions):
                isopathlist.append(os.path.join(root, file))
                isofilelist.append(file)
                isocount += 1

    print(outputsep)

    if not isofilelist:
        print("No ISOs found.")
        print(outputsep)
        print("Program terminated.")
        return

    print("Number of ISOs: {}".format(isocount))
    print()
    print(outputsep)
    print("Output File: {}".format(os.path.abspath(outfile)))
    print(outputsep)

    while True:
        inp = input("q to quit, s to show all isos, c to continue: ").lower()
        if inp == 'q':
            print(outputsep)
            print("Program terminated.")
            return
        elif inp == 's':
            for i in isofilelist:
                print(i, end=', ')
        elif inp == 'c':
            break
        print('\n' + outputsep)

    starttime = time.time()
    print(outputsep)
    errorlist = parprocess(isocount, isopathlist, isofilelist, outfile, starttime)
    print(outputsep)
    if errorlist:
        print("Errors: ")
        print(outputsep)
        for i in errorlist:
            print(i, end=', ')
        print('\n' + outputsep)
        print("{} Errors".format(len(errorlist)))
    else:
        print("No Errors")
    print(outputsep)
    print("Output File Size: {} Byte".format(os.path.getsize(outfile)))
    print(outputsep)
    print("Program completed.")
    winsound.PlaySound("*", winsound.SND_ALIAS)


def finder():
    outfile = "exes.json"
    outputsep = "--------------------"

    print("Loading JSON...")
    try:
        with open(outfile) as file:
            ordict = json.load(file)
    except:
        print("File not found or empty.")
        return
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


if __name__ == "__main__":

    # On Windows calling this function is necessary.
    multiprocessing.freeze_support()

    # Module multiprocessing is organized differently in Python 3.4+
    try:
        # Python 3.4+
        if sys.platform.startswith('win'):
            import multiprocessing.popen_spawn_win32 as forking
        else:
            import multiprocessing.popen_fork as forking
    except ImportError:
        import multiprocessing.forking as forking

    if sys.platform.startswith('win'):
        # First define a modified version of Popen.
        class _Popen(forking.Popen):
            def __init__(self, *args, **kw):
                if hasattr(sys, 'frozen'):
                    # We have to set original _MEIPASS2 value from sys._MEIPASS
                    # to get --onefile mode working.
                    os.putenv('_MEIPASS2', sys._MEIPASS)
                try:
                    super(_Popen, self).__init__(*args, **kw)
                finally:
                    if hasattr(sys, 'frozen'):
                        # On some platforms (e.g. AIX) 'os.unsetenv()' is not
                        # available. In those cases we cannot delete the variable
                        # but only set it to the empty string. The bootloader
                        # can handle this case.
                        if hasattr(os, 'unsetenv'):
                            os.unsetenv('_MEIPASS2')
                        else:
                            os.putenv('_MEIPASS2', '')


        # Second override 'Popen' class with our modified version.
        forking.Popen = _Popen

    while True:
        modeselect = input("Select mode(i for indexer, f for finder, q to quit): ")
        if modeselect == 'i':
            indexer()
        if modeselect == 'f':
            finder()
        if modeselect == 'q':
            sys.exit()
