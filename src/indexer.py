from isoarchive import *


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
        input("Press Enter to terminate Program...")
        sys.exit()

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
            sys.exit()
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
        errorliststr = ""
        for i in errorlist:
            errorliststr += i + ', '
        print(errorliststr)
        print('\n' + outputsep)
        print("{} Errors".format(len(errorlist)))
    else:
        print("No Errors")
    print(outputsep)
    print("Output File Size: {} Byte".format(os.path.getsize(outfile)))
    print(outputsep)
    print("Program completed.")
    winsound.PlaySound("*", winsound.SND_ALIAS)
    input("Press Enter to terminate Program...")
