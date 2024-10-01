' --------------------------------- 
'      EIC Data Analysis Script 
'      By Michael Brigham, bs16mwb@leeds.ac.uk / michaelwbrigham@gmail.com 
'      September 2023
' --------------------------------- 
'
'--------------------
'--- DEFAULT ARGS ---
'--------------------
'
' Default paths for input boxes (otherwise will be blank):
Dim csv_file_path 
csv_file_path = "C:\Michael\Orderofevents_July24\Scripts\chloromeotrp_metabolites.csv"

Dim export_dir_path
export_dir_path = "C:\Michael\Orderofevents_July24\Scripts\"
'
' Default ppm tolerance:
Dim ppm_tol
ppm_tol = 20
'
'-----------------
'--- FUNCTIONS ---
'-----------------
'
' csv_to_mzdictarr
'   read in a csv file containing rows of eic. each row contains the eic name followed by a list of mz corresponding to the eic
'   e.g. [m+h]+ , [m+na]+, [m+h-nh3]+
'
'   -- ARGS ----
'   path_to_file (str) - path to the eic csv file
Function csv_to_mzdictarr(path_to_file)
    ' file reading stuff
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set obj_open = obj_FSO.OpenTextFile(path_to_file, 1)
    Dim arr_file_name()

    file_content = obj_open.ReadAll ' read file
    file_content_arr = Split(file_content, VbCrLF) ' split the file into lines (VbCrLF => new line)
    obj_open.Close ' close the file
    Set obj_open = Nothing ' set to null
    Set obj_FSO = Nothing ' set to null

    ' going over the file line array
    Redim eic_obj_arr(CInt(UBound(file_content_arr) - 1)) ' create eic array and set to number of rows (minus 1 ignores the header)
    
    For line_counter = 1 to UBound(file_content_arr) ' go over each line
        Dim split_line
        split_line = Split(file_content_arr(line_counter), ",") ' split it up into a list by each commas i.e. "A,B,C" => ["A", "B", "C"]
        Dim filtered_split_line
        'filtered_split_line = Filter(split_line, "", False) ' remove any blank values
        filtered_split_line = split_line
        Dim mz_value
        mz_values = Array(UBound(filtered_split_line) - 1) ' create an array of a size equal to the split line arr minus one (to ignore the mz name)
        
        Dim i
        i = 1
        For i = 1 to UBound(filtered_split_line) ' create a list of the mz values (starting at index 1 to ignore the mz name)
            mz_values(i-1) = filtered_split_line(i)
            i = i+1 
        Next

        Dim obj_dictionary ' create a dictionary
        Set obj_dictionary = CreateObject("Scripting.Dictionary")
        obj_dictionary.CompareMode = vbTextCompare
        obj_dictionary.Add "eic_name", filtered_split_line(0) ' set eic name to the first item in the line array
        obj_dictionary.Add "eic_mz", mz_values ' set mz values to the mz value arr

        Set eic_obj_arr(line_counter - 1) = obj_dictionary ' assign this dict to a position in the eic obj arr
    Next
    
    csv_to_mzdictarr = eic_obj_arr ' make eic_dic_arr the output
    
End Function

' stdmz_to_eic
'   convert an mz value corresponding to a standard to an eic dict object.
'
'   -- ARGS ----
'   mz_val (num) - m/z of the standard

Function stdmz_to_eic(mz_val)
    Dim obj_dictionary
    Set obj_dictionary = CreateObject("Scripting.Dictionary")
    obj_dictionary.CompareMode = vbTextCompare
    obj_dictionary.Add "eic_name", "standard"
    obj_dictionary.Add "eic_mz", mz_val
    stdmz_to_eic = obj_dictionary
End Function

' calculate_ppm_tolerance
'   within the mz_dict take the a list of eic_mz values and work out the tolerance based on 
'   the ppm of the largest value within this list.
'
'   -- ARGS ----
'   ppm_tolerance (num) - ppm
'   mz_dict (dict obj) - an eic mz dict object. this function will look at the eic_mz arr.
Function calculate_ppm_tolerance(ppm_tolerance, mz_dict) 
    Dim mz_values
    If mz_dict.Exists("eic_mz") Then mz_values = mz_dict("eic_mz")

    Dim highest_mz  
    highest_mz = mz_values(0) 
    For i = 1 To UBound(mz_values) 
        If eic_obj.mz_arr(i) > highest_mz Then 
            highest_mz = eic_obj.mz_arr(i) 
        End If 
    Next 
         
    Dim mass_tolerance 
    mass_tolerance = ppm_tolerance * highest_mz / 10^6

    calculate_ppm_tolerance = mass_tolerance
End Function

' create_query_str
'   create the query string that is utilised by the bruker EIC obj. converts the list of eic_mz 
'   values to a string in which each values is separated by a colon.
'
'   -- ARGS ----
'   mz_dict (dict obj) - an eic mz dict object. this function will look at the eic_mz arr.
Function create_query_str(mz_dict)
    Dim mz_values
    If mz_dict.Exists("eic_mz") Then mz_values = mz_dict("eic_mz")

    Dim range_query 
    range_query = "" 

    Dim mz 
    For Each mz in mz_values ' For each value in the mz array... 
        range_query = range_query + Cstr(mz) + ";" ' Add this the the current query string 
    Next 
    range_query = Mid(range_query, 1, Len(range_query) - 1) ' Remove the final ; from the query string as it should not be there

    create_query_str = range_query
End Function

'------------
'--- MAIN ---
'------------

' get the dir to the csv file
Dim file_name 
file_name=InputBox("Enter the dir to the mz csv file:",csv_file_path,csv_file_path)

Dim output_dir 
output_dir=InputBox("Enter the dir for the output file:",export_dir_path,export_dir_path)

Dim mz_list
mz_list = csv_to_mzdictarr(file_name) ' generate mz_list from csv file

Dim currentAnalysis 
For Each currentAnalysis in Application.Analyses 
    currentAnalysis.Chromatograms.Clear 
    
    Dim mz_dict  
    For Each mz_dict in mz_list 
        Dim query 
        query = create_query_str(mz_dict)

        Dim tolerance
        tolerance = calculate_ppm_tolerance(ppm_tol, mz_dict) 

        Set EIC = CreateObject("DataAnalysis.EICChromatogramDefinition") ' Create an EIC 
 
        EIC.Range = query
        EIC.WidthLeft = tolerance 
        EIC.WidthRight = tolerance 

        currentAnalysis.Chromatograms.AddChromatogram EIC ' Add the chromatogram 
    Next 


    Set TIC = CreateObject("DataAnalysis.TICChromatogramDefinition") 
    currentAnalysis.Chromatograms.AddChromatogram TIC 
 
    Set BPC = CreateObject("DataAnalysis.BPCChromatogramDefinition") 
    currentAnalysis.Chromatograms.AddChromatogram BPC 
     
    ' EXPORT DATA  
 
    Dim output_file_path 
    output_file_path = output_dir + "\EICs_"+currentAnalysis.Name+".tsv" 'path that the file will be saved to 
    
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set file = obj_FSO.CreateTextFile(output_file_path,true) 'true will overwrite existing files 
 
    'iterate over all chromatograms  
    Dim chrom_counter 
    chrom_counter = 0 
    Dim upper_bound 
    upper_bound = UBound(mz_list)+2 ' number corresponding to total number of eics from the mz 
                                    ' the tic and the bpc
    For Each chrom in currentAnalysis.Chromatograms 
        'set the arrays declared above with the data from the current chromatogram 
        chrom.ChromatogramData rt, intensity 
         
        ' write each data point to the file. !something! is a notation used for parsing headers within the python script. 
        If chrom_counter < upper_bound - 1 Then ' if the chromo is from the eic csv then...
            file.WriteLine("!NAME!" + mz_list(chrom_counter)("eic_name") + "!CHROMNAME!" + chrom.Name)
        ElseIf chrom_counter = upper_bound - 1 Then ' if the chromo is the tic 
            file.WriteLine("!NAME!" + "TIC" + "!CHROMNAME!" + chrom.Name) 
        ElseIf chrom_counter = upper_bound Then ' if the chromo is the bpc
            file.WriteLine("!NAME!" + "BPC" + "!CHROMNAME!" + chrom.Name)
        End If

        Dim i 
        For i = 0 to UBound(rt) ' then loop over data points 
            file.WriteLine(rt(i) & vbTab & intensity(i)) ' and add rt and intensity tab separated 
        Next 
        chrom_counter = chrom_counter + 1
    Next
Next