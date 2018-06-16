option explicit

sub writeFile(filename as string)

    open   filename for output as #1
    print# 1, "This is the file"
    print# 1, "with the filename  " & filename
    close# 1

end sub


sub main()

    dim file_1 as string
    dim file_2 as string

    file_1 = environ("TEMP") & "\one.txt"
    file_2 = environ("TEMP") & "\two.txt"

    call writeFile(file_1)
    call writeFile(file_2)

    call ShellExecute(0, "Open", file_1       , ""    , "", 1)
    call ShellExecute(0, "Open", "notepad.exe", file_2, "", 1)

end sub
