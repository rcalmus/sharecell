% Uses ActiveX to put a formula "=SUM(A1..A100)" into a cell in an Excel workbook.  Useful if you want the formula to depend on how much data you poked into the worksheet.
clc;    % Clear the command window.
close all;  % Close all figures (except those of imtool.)
clear;  % Erase all existing variables. Or clearvars if you want.
workspace;  % Make sure the workspace panel is showing.
format long g;
format compact;
fontSize = 14;

% Launch an Excel server using ActiveX (Windows ONLY).
excelObject = actxserver('Excel.Application');
% Create the filename of the existing workbook.
fullFileName = fullfile(pwd, 'Example.xlsx');
% Make sure the file exists.
if ~isfile(fullFileName)
	errorMessage = sprintf('The workbook file does not exist:\n%s', fullFileName);
	uiwait(errordlg(errorMessage));
	return;
end
% Open the workbook from disk.
excelWorkbook = excelObject.workbooks.Open(fullFileName);
% Excel is invisible so far.  Make it visible.
excelObject.Visible = true;
% Create a string with the formula just like you'd have it in Excel.
yourFormula = '=SUM(A1..A100)'; % No spaces allowed.
% Assign the formula to the cell "B1".
excelWorkbook.ActiveSheet.Range('B1').Formula = yourFormula;
% Save the current state of the workbook.
excelWorkbook.Save;
% Close the workbook.  Excel will stay open but be hidden.
% You can still see it as "Microsoft Excel" in Task Manager.
excelWorkbook.Close;
% Shut down the Excel server instance.
excelObject.Quit;
% Even after quitting, you can still see it as "Microsoft Excel" in Task Manager.
% Clear the excel object variable from MATLAB's memory.
clear('excelObject', 'excelWorkbook', 'yourFormula');
% The clear finally shuts down the server and it no longer appears in Task Manager.
fprintf('Done interacting with Excel.\n');