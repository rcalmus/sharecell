% Demo macro to use ActiveX to write numerical arrays, cell arrays, and figure windows to two different worksheets in an Excel workbook file.
% IMPORTANTE NOTE:  Uses xlswrite1, available from the File Exchange, so BE SURE TO DOWNLOAD THAT!!!
% http://www.mathworks.com/matlabcentral/fileexchange/10465-xlswrite1
function ExcelDemo
clc;
workspace;  % Make sure the workspace panel is showing.
try
	if	~contains(computer, 'WIN')
		% If it's not a Windows computer, they can't use ActiveX, so just exit.
		msgbox('ActiveX only works with Windows computers.');
		return;
	end
	% Check for existence of the xlswrite1() function.
	xlswrite1FunctionPath = which('xlswrite1');
	if isempty(xlswrite1FunctionPath)
		errorMessage = sprintf('The xlswrite1() function was not found on the search path.\nPlease download it from \nhttp://www.mathworks.com/matlabcentral/fileexchange/10465-xlswrite1\n and save it to a folder on your path.\nClick OK to go there.');
		uiwait(errordlg(errorMessage));
		% Open their default web browser to this page on the MATLAB Central File Exchange.
		web('http://www.mathworks.com/matlabcentral/fileexchange/10465-xlswrite1')
		return; % Can't do anything further, so just exit.
	end
	
	promptMessage = sprintf('First you will specify the name of a brand new Excel file that we will create.');
	titleBarCaption = 'Continue?';
	button = questdlg(promptMessage, titleBarCaption, 'Continue', 'Cancel', 'Continue');
	if strcmpi(button, 'Cancel')
		return;
	end
	
	fullFileName = GetExcelWorkbookFileName();
	if isempty(fullFileName)
		% User clicked Cancel.
		return;
	end
	try
		% See if there is an existing instance of Excel running.
		% If Excel is NOT running, this will throw an error and send us to the catch block below.
		Excel = actxGetRunningServer('Excel.Application');
		% If there was no error, then we were able to connect to it.
	catch
		% No instance of Excel is currently running.  Create a new one.
		% Normally you'll get here (because Excel is not usually running when you run this).
		Excel = actxserver('Excel.Application');
	end
	
	% Prepare proper filename extension.
	% Get the Excel version because if it's version 11 (Excel 2003) the file extension should be .xls,
	% but if later than version 12.0 (Excel 2007) then we'll need to use an extension of .xlsx to avoid nag messages.
	excelVersion = str2double(Excel.Version);
	if excelVersion < 12
		excelExtension = '.xls';
	else
		excelExtension = '.xlsx';
	end
	% If it's a brand new file (not yet existing) make sure it has the latest extension,
	% regardless of what extension they gave it in the file open dialog box.
	if ~exist(fullFileName, 'file')
		% Parse what they actually specified.
		[folder, baseFileName, extension] = fileparts(fullFileName);
		% Build the new filename with the proper extension.
		fullFileName = fullfile(folder, [baseFileName, excelExtension]);
	end
	
	% Determine the proper format to save the files in.  It depends on the extension (Excel version).
	switch excelExtension
		case '.xls' %xlExcel8 or xlWorkbookNormal
			xlFormat = -4143;
		case '.xlsb' %xlExcel12
			xlFormat = 50;
		case '.xlsx' %xlOpenXMLWorkbook
			xlFormat = 51;
		case '.xlsm' %xlOpenXMLWorkbookMacroEnabled
			xlFormat = 52;
		otherwise
			xlFormat = -4143;
	end
	
	%=====================================================================================================================
	% Part 1: CREATE A NEW WORKBOOK AND WRITE STUFF INTO IT.  DO PARTS 1a AND 1b.
	Excel.visible = true; % Make Excel appear so we can see it, otherwise it is hidden.
	% Part 1a:
	% Right now, Excel is open but there is no workbook in it. (You'll see this if you maximize Excel.)
	% Add a new workbook.
	ExcelWorkbook = Excel.workbooks.Add;
	% It will have now have one sheet in it, called Sheet1, which is what Excel calls newly created sheets by default.
	% Let's rename the sheet 'My Magic Data'.
	ExcelWorkbook.ActiveSheet.Name = 'My Magic Data';
	% Tell it to not wait and pop up alerts like "This file exists.  Do you want to overwrite it."
	Excel.DisplayAlerts = false;
	% Save this workbook we just created to disk.
	ExcelWorkbook.SaveAs(fullFileName, xlFormat);
	% Close the workbook.  This will show you how you can open an existing workbook if Excel is already open.
	ExcelWorkbook.Close(false);
	
	% Open the workbook we just created and closed.
	% Part 1b:
	% This will also open up an existing workbook file if one exists with that filename
	% even if you didn't create a brand new one and close it.
	% IMPORTANT NOTE: If you're ONLY opening up an existing workbook and NOT creating a new one,
	% then you can skip Part 1a and start right here with Part 1b.
	invoke(Excel.Workbooks, 'Open', fullFileName);
	ExcelWorkbook = Excel.ActiveWorkbook; % Get the workbooks object by itself for convenient reference later.
	
	%----------------------------------------------------------------------------------------------------------------
	% Create some sample data.
	myMagicData = magic(12);
	myOtherData = 50 * rand(1, 10);
	%----------------------------------------------------------------------------------------------------------------
	
	% Then run the new xlswrite1 function as many times as needed or in a loop
	% (for example xlswrite1(fullFileName, yourArrayName, XL_CellLocation).
	% IMPORTANT NOTE: the Excel object variable MUST exist in the routine that calls xlswrite1()
	% and it MUST be named "Excel" EXACTLY because xlswrite1() has this line it it:
	% 	Excel = evalin('caller', 'Excel');
	
	% 1: Write the magic number array variable to the sheet.
	xlswrite1(fullFileName, myMagicData, 'My Magic Data', 'B2');
	
	% 2: Create a hard-coded cell array for the column headers then write them to the workbook.
	caColHeader = {'Column Header 1', 'Column Header 2', 'Column Header 3', 'Column Header 4'};
	xlswrite1(fullFileName, caColHeader, 'My Magic Data', 'B1');
	
	% 3: Create a cell array dynamically for the row headers then write them to the workbook.
	% IMPORTANT NOTE: to write strings into a single cell you must have the string inside a cell or cell array.
	% Otherwise each character of the string will show up in a separate worksheet cell in Excel instead of a single cell.
	numberOfRows = size(myMagicData, 1); % number of rows in myData.
	caRowHeader = cell(numberOfRows, 1); % Create a column vector because we want the headers to be in a column, not a row.
	for row = 1 : numberOfRows % Create a header for every row of our data.
		caRowHeader{row} = sprintf('Row Header %d', row);
	end
	% Poke onto the sheet with name 'My Magic Data'.
	xlswrite1(fullFileName, caRowHeader, 'My Magic Data', 'A2');
	
	% 4: Write the random number array variable into a different sheet.
	xlswrite1(fullFileName, myOtherData, 'My Other Data', 'B2');
	
	% 5: Just for fun, add comments to cells A1:12 on sheet #1.
	worksheets = Excel.sheets;
	thisSheet = get(worksheets, 'Item', 1);
	for k = 1 : numberOfRows
		myComment = sprintf('Comment for cell A%d', k);
		cellReference = sprintf('A%d', k);
		theCell = thisSheet.Range(cellReference);
		theCell.AddComment(myComment);
	end

	% 6: Demonstrate copy and paste between worksheets, or within a worksheet
	% Activate sheet #1.
	sheetIndex = 1; % The 'My Magic Data' sheet.
	ExcelWorkbook.Sheets.Item(sheetIndex).Activate;
	sourceCellRange = 'B2..M13';
	Excel.Range(sourceCellRange).Copy; % Copy selected cells to clipboard.
	% 	Excel.ActiveSheet.Cells.Copy; % Copy entire worksheet to clipboard
	% Copy this block of cells to sheet 2.  (You can skip this if you want it to stay in sheet 1.)
	sheetIndex = 2; % The 'My Magic Data' sheet.
	ExcelWorkbook.Sheets.Item(sheetIndex).Activate; % Activate destination sheet.
	destinationCellRange = 'B13'; % Upper left corner.
	Excel.Range(destinationCellRange).Select; % Select destination cell (upper left corner).
	% Paste the worksheets from the source worksheet into the destination worksheet.
	Excel.ActiveSheet.Paste;
	% At this point the whole range of pasted cells is selected.
	% This is dangerous, should the user accidentally hit delete.
	% Put "cursor" or active cell at A1, the upper left cell so that only one cell is highlighted.
	Excel.Range('A1').Select;
	
	% 7: Create a plot and paste it into Excel.  Need to do this by creating a figure and saving the figure
	% to a disk file, then importing the disk file into the specified Excel worksheet.
	PasteFigureIntoExcel(Excel);
	
	% Then save the workbook.  Here we are using the invoke function, (just to show you anohter way to do it)
	% but you could use the ActiveX method as well.
	invoke(Excel.ActiveWorkbook,'Save');
	
	% Then run the following code to close the activex server (shut down Excel.)
	% Delete or comment out the next three lines if you want to leave Excel open.
	Excel.Quit;
	Excel.delete;
	clear Excel;
	writeMessage = sprintf('Done!\nThis Excel workbook has been created:\n%s', fullFileName);
	uiwait(msgbox(writeMessage));
	
	%======================================================================================================================
	% Part 2: OPEN AN EXISTING WORKBOOK AND READ STUFF FROM IT.
	promptMessage = sprintf('Now I am going to open that Excel workbook we just created\n\n%s\n\nand read from it.\nClick OK to continue and open an Excel instance and read this data from the workbook.\nClick Exit to exit this function', fullFileName);
	button = questdlg(promptMessage, 'Open existing workbook', 'OK', 'Exit', 'OK');
	drawnow;	% Refresh screen to get rid of dialog box remnants.
	if strcmpi(button, 'Exit')
		return;
	end
	if ~isfile(fullFileName)
		warningMessage = sprintf('Error: the file:\n%s\ndoes not exist!', fullFileName);
		WarnUser(warningMessage);
		return;
	end
	
	% Start up Excel as an ActiveX server.
	try
		% See if there is an existing instance of Excel running.
		% If Excel is NOT running, this will throw an error and send us to the catch block below.
		Excel = actxGetRunningServer('Excel.Application');
		% If there was no error, then we were able to connect to it.
	catch
		% No instance of Excel is currently running.  Create a new one.
		% Normally you'll get here (because Excel is not usually running when you run this).
		Excel = actxserver('Excel.Application');
	end
	Excel.visible = true;
	% Tell it to not ask you for confirmation when resaving or closing the workbooks.
	Excel.DisplayAlerts = false;
	% Prevent beeps from sounding if we try to do something Excel doesn't like.
	% 	Excel.EnableSound = false;
	% Open the existing workbook named in the variable fullFileName.
	% 	invoke(Excel.Workbooks,'Open', fullFileName);
	% Just to show you that there is a method that is equivalent to doing it via the invoke function,
	% like we used before, this time we will open it with the Excel.Workbooks.Open() method.
	Excel.Workbooks.Open(fullFileName); % Could also use Workbook.Open(fullFileName) if we want.
	% Get handles to the Workbook and to the Worksheets collection, for convenience.
	Workbook = Excel.ActiveWorkbook;
	Worksheets = Workbook.sheets;
	% Get the number of worksheets in the source workbook. (Just for fun, so you can see we can get it.)
	numberOfSourceSheets = Worksheets.Count;
	
	% First let's go to My Magic Data and pull out the magic number array.
	sheetIndex = 1; % The 'My Magic Data' sheet.
	Worksheets.Item(sheetIndex).Activate;
	% Get cells from range 'B2..M13'.  Results go into a cell array caExcelCellInfo.
	% Use MATLAB's general purpose get() function.
	caExcelCellInfo = get(Excel.ActiveSheet, 'Range', 'B2', 'M13');
	% Extract the numbers and convert to a double array.
	% Values displayed in the command window because I'm leaving the semicolon off the line of code.
	myDataRetrieved = cell2mat(caExcelCellInfo.value)
	
	% Next let's go to My Magic Data and pull out the column header array.
	% Get cells from range 'B1..F1'.  Results go into a cell array caExcelCellInfo.
	caExcelCellInfo = get(Excel.ActiveSheet, 'Range', 'B1', 'E1');
	% Extract the numbers and convert to a cell array of strings.
	% Values displayed in the command window because I'm leaving the semicolon off the line of code.
	caColHeaderRetrieved = caExcelCellInfo.value
	
	% Next let's go to My Magic Data and pull out the row header array.
	% Get cells from range 'A2..A13'.  Results go into a cell array caExcelCellInfo.
	caExcelCellInfo = get(Excel.ActiveSheet, 'Range', 'A2', 'A13');
	% Extract the numbers and convert to a cell array of strings.
	% Values displayed in the command window because I'm leaving the semicolon off the line of code.
	caRowHeaderRetrieved = caExcelCellInfo.value
	
	% Next let's go to myOtherSheetName and pull out the number arrays.
	% First array: a row vector in row 2
	sheetIndex = 2; % The 'myOtherSheetName' sheet.
	Worksheets.Item(sheetIndex).Activate;
	% Get cells from range 'B2..K2'.  Results go into a cell array caExcelCellInfo.
	caExcelCellInfo = get(Excel.ActiveSheet, 'Range', 'B2', 'K2');
	% Extract the numbers and convert to a double array.
	% Values displayed in the command window because I'm leaving the semicolon off the line of code.
	myOtherDataRetrieved_A = cell2mat(caExcelCellInfo.value)
	% Second array: a matrix covering cells from B13 to M24.
	% Get cells from range 'B2..K2'.  Results go into a cell array caExcelCellInfo.
	caExcelCellInfo = get(Excel.ActiveSheet, 'Range', 'B13', 'M24');
	% Extract the numbers and convert to a double array.
	% Values displayed in the command window because I'm leaving the semicolon off the line of code.
	myOtherDataRetrieved_B = cell2mat(caExcelCellInfo.value)
	
	% Then run the following code to close the activex server (shut down Excel.)
	% Delete or comment out the next three lines if you want to leave Excel open.
	Excel.Quit;
	Excel.delete;
	clear Excel;

	% See if user wants to open Excel and the workbook.
	promptMessage = sprintf('Look in the command window to see the retrieved cell values.\nDo you want to open the workbook in Excel?');
	titleBarCaption = 'Open?';
	buttonText = questdlg(promptMessage, titleBarCaption, 'Yes - Open It', 'No', 'Yes - Open It');
	if contains(buttonText, 'Yes', 'IgnoreCase', true)
		% Launch Excel and open the workbook we just created.
		winopen(fullFileName);
	end

catch ME
	errorMessage = sprintf('Error in function %s() at line %d.\n\nError Message:\n%s\n\nDo you have the file already open in Excel?\nIf so, please close it before you run this demo.', ...
		ME.stack(1).name, ME.stack(1).line, ME.message);
	WarnUser(errorMessage);
	if exist('Excel', 'var')
		% Shut down Excel.
		Excel.Quit;
		Excel.delete;
		clear Excel;
	end
end
return;

% End of main function: ExcelDemo.m -----------------------------

%==========================================================================================================================
% Gets the name of the workbook from the user.
function fullExcelFileName = GetExcelWorkbookFileName()
try
	fullExcelFileName = [];  % Default.
	% Ask user for a filename.
	FilterSpec = {'*.xls;*.xlsx', 'Excel workbooks (*.xls)'; '*.*', 'All Files (*.*)'};
	DialogTitle = 'Save workbook file name';
	% Get the default filename.  Make sure it's in the folder where this m-file lives.
	% (If they run this file but the cd is another folder then pwd will show that folder, not this one.
	thisFile = mfilename('fullpath');
	[thisFolder, baseFileName, extension] = fileparts(thisFile);
	DefaultName = sprintf('%s/%s.xls', thisFolder, baseFileName);
	[fileName, specifiedFolder] = uiputfile(FilterSpec, DialogTitle, DefaultName);
	if fileName == 0
		% User clicked Cancel.
		return;
	end
	% Parse what they actually specified.
	[folder, baseFileName, extension] = fileparts(fileName);
	% Create the full filename, making sure it has a xls filename.
	if strfind(extension, '.xls')
		% It has a standard Excel extension of xls or xlsx.
		fullExcelFileName = fullfile(specifiedFolder, [baseFileName extension]);
	else
		% They gave it some other extension.
		fullExcelFileName = fullfile(specifiedFolder, [baseFileName '.xls']);
		% We can convert it to .xlsx extension if we determine later that they have a newer version of Excel.
	end
	
	if exist(fullExcelFileName, 'file')
		promptMessage = sprintf('You specified the name of an existing Excel file.\n%s\nCan I delete it?', fullExcelFileName);
		titleBarCaption = 'Delete File?';
		button = questdlg(promptMessage, titleBarCaption, 'Yes - delete it', 'No - keep it', 'No - keep it');
		if strcmpi(button, 'Yes - delete it')
			% Delete the existing file.
			delete(fullExcelFileName);
		else
			fullExcelFileName = []; % Signal that they want to quit.
			return;
		end
	end
catch ME
	errorMessage = sprintf('Error in function %s() at line %d.\nDo you have the file already open in Excel?\nPlease check that the file is not open anywhere.\n\nError Message:\n%s', ...
		ME.stack(1).name, ME.stack(1).line, ME.message);
	WarnUser(errorMessage);
end

%==========================================================================================================================
% DeleteEmptyExcelSheets: deletes all empty sheets in the active workbook.
% This function looped through all sheets and deletes those sheets that are
% empty. Can be used to clean a newly created xls-file after all results
% have been saved in it.
function DeleteEmptyExcelSheets(excelObject)
try
	% 	excelObject = actxserver('Excel.Application');
	% 	excelWorkbook = excelObject.workbooks.Open(fileName);
	worksheets = excelObject.sheets;
	sheetIdx = 1;
	sheetIdx2 = 1;
	numSheets = worksheets.Count;
	% Prevent beeps from sounding if we try to delete a non-empty worksheet.
	originalSoundSetting = excelObject.EnableSound;
	excelObject.EnableSound = false;
	
	% Loop over all sheets
	while sheetIdx2 <= numSheets
		% Saves the current number of sheets in the workbook
		temp = worksheets.count;
		% Check whether the current worksheet is the last one. As there always
		% need to be at least one worksheet in an xls-file the last sheet must
		% not be deleted.
		if or(sheetIdx>1,numSheets-sheetIdx2>0)
			% worksheets.Item(sheetIdx).UsedRange.Count is the number of used cells.
			% This will be 1 for an empty sheet.  It may also be one for certain other
			% cases but in those cases, it will beep and not actually delete the sheet.
			if worksheets.Item(sheetIdx).UsedRange.Count == 1
				worksheets.Item(sheetIdx).Delete;
			end
		end
		% Check whether the number of sheets has changed. If this is not the
		% case the counter "sheetIdx" is increased by one.
		if temp == worksheets.count
			sheetIdx = sheetIdx + 1;
		end
		sheetIdx2 = sheetIdx2 + 1; % prevent endless loop...
	end
	excelObject.EnableSound = originalSoundSetting;
catch ME
	errorMessage = sprintf('Error in function %s() at line %d.\n\nError Message:\n%s', ...
		ME.stack(1).name, ME.stack(1).line, ME.message);
	WarnUser(errorMessage);
end
return;

%==========================================================================================================================
% Create a figure window with a plot (axes control) on it and paste it into the Excel worksheet.
% Adapted from https://www.mathworks.com/matlabcentral/answers/91547-how-do-i-insert-a-matlab-figure-into-an-excel-workbook-through-activex
function PasteFigureIntoExcel(Excel)
try
	% Create sample image from figure	
	hFig = figure; % Bring up a new, separate figure.
	period = pi;
	x = linspace(-2*pi, 2*pi, 400);
	mySineWave = 10 * cos(2 * pi * x / period);
	% Plot it.
	plot(x, mySineWave, 'b-', 'LineWidth', 2);
	xlabel('x', 'FontSize', 20);
	ylabel('y', 'FontSize', 20);
	title('My Sine Wave', 'FontSize', 20);
	grid on;
	
	% Create a filename to save the figure window to disk for temporarily.
	tempImageFullFileName = fullfile(pwd, 'Delete_me.png');	
	print('-dpng', tempImageFullFileName);	% Save it to disk.
	uiwait(msgbox('We saved this plot to disk and will paste it into the workbook.'));
	close(hFig); % Close down figure because we don't need it anymore.
	% Alternative 1 BEGIN.	
	% Get handle to Excel COM Server	
% 	Excel = actxserver('Excel.Application');	
% 	% Set it to visible	
% 	set(Excel,'Visible',1);	
	
	% Optionally, add a Workbook	
% 	Workbooks = Excel.Workbooks;	
% 	Workbook = invoke(Workbooks, 'Add');	
	
	% Get a handle to Sheets and select Sheet 1	
% 	Sheets = Excel.ActiveWorkBook.Sheets;	
% 	Sheet1 = get(Sheets, 'Item', 1);	
% 	Sheet1.Activate;
	% Alternative 1 END.
	
	% Alternative 2 BEGIN.	
	% Use whatever sheet is active at the moment.  This should be setup in advance of calling this function.
	currentSheet = Excel.ActiveSheet;
	% Alternative 2 END.	
	
	% Alternative 1 BEGIN.	
	% Get a handle to Shapes for Sheet 1	
	Shapes = currentSheet.Shapes;	
	% Add image	by importing one from an image file on disk.  Place at a certain position.  Last 4 arguments are : x, y, width, height.
	Shapes.AddPicture(tempImageFullFileName, 0, 1, 200, 40, 300, 235);	
	% Alternative 1 END.
	
	% Alternative 2 BEGIN.	
	% Add image	
% 	Sheet1.invoke('Pictures').Insert(tempImageFullFileName);	
	% Alternative 2 END.	
	
	% Save the workbook and Close Excel	
% 	invoke(Workbook, 'SaveAs', fullfile(pwd, '\myfile.xls'));	
% 	invoke(Excel, 'Quit');

	% Delete the temporary image file.
	delete(tempImageFullFileName);
catch ME
	errorMessage = sprintf('Error in function %s() at line %d.\n\nError Message:\n%s', ...
		ME.stack(1).name, ME.stack(1).line, ME.message);
	WarnUser(errorMessage);
end
return;

%==========================================================================================================================
function WarnUser(warningMessage)
uiwait(warndlg(warningMessage));
return; % from WarnUser()


% Snippet to run a macro stored inside the Excel workbook file:
% Excel=actxserver('Excel.Application');
% excelFullFileName = fullfile(pwd, '\Document.xls');
% eW = Excel.Workbooks;
% eF = eW.Open(excelFullFileName);
% invoke(Excel,'Run','Document.xls!My_Macro');
% eF.Save;
% invoke(Excel, 'Quit');


