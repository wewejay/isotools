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
        resultstr = ""
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
            line = "{} | {} | {}\n".format(indexstr, filenamestr, isonamestr)
            resultstr += line
        print(resultstr)
        print("{} Results.".format(len(result)))


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
