function excelFileNames = rawDataToExcel(rawData)

    addpath('.\excelFormattingTools\');
    
    fields = fieldnames(rawData);
    for iPaperPart = 1:length(fields)
        fprintf('Currently writing data from manuscript section "%s"\n',upper(fields{iPaperPart}));
        excelFileNames{iPaperPart} = dataToExcel(fields{iPaperPart},rawData.(fields{iPaperPart}));
        fprintf('\n');
    end
end

function excelFileName = dataToExcel(paperSection,data)

    global Excel;
    
    cellOffset = @(x) sprintf('B%d',x);
    
    excelFileName = fullfile('.',[paperSection,'_data.xlsx']);
    excelFileNameAbs = fullfile(cd,[paperSection,'_data.xlsx']);
    
    %     excelVersion = str2double(Excel.Version);
    % 	if excelVersion < 12
    % 		excelExtension = '.xls';
    % 	else
    % 		excelExtension = '.xlsx';
    % 	end
    
    figureNames = sort(fieldnames(data));
    try
        sheetNames = cellfun(@(x) strrep([upper(x(1)),lower(x(2:end))],'_',' '),figureNames,'UniformOutput',false);
    catch
        sheetNames = cellfun(@(x) upper(strrep(x,'_',' ')),figureNames,'UniformOutput',false);
    end
    
    writecell({' '},excelFileName,'Sheet',sheetNames{1},'Range','A1:A1');
    
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
    invoke(Excel.Workbooks, 'Open', excelFileNameAbs);
    
    for iFigure = 1:length(figureNames)
    
        curOffset = 1;
    
        fprintf('Currently writing figure "%s"\n',figureNames{iFigure});

        FoundFields = StructFind(data.(figureNames{iFigure}),'tables');
        pathsToTables = cellfun(@(x)strcat(figureNames{iFigure},x),FoundFields,'UniformOutput',false)';
        if isempty(pathsToTables)
            pathsToTables = figureNames(iFigure);
        end
        pathsToTables = cellfun(@(x)strcat(x,'.tables'),pathsToTables,'UniformOutput',false);
        for iTableSet = 1:length(pathsToTables)
            curSet = pathsToTables{iTableSet};
    
            fprintf('Currently writing table collection "%s"\n',curSet);
    
            curFieldPath = split(curSet,'.');
            refStruct = struct();
            for iRef = 1:length(curFieldPath)
                refStruct(iRef).type = '.';
                refStruct(iRef).subs = curFieldPath{iRef};
            end
            curTables = subsref(data,refStruct);
            tableNames = fieldnames(curTables);
    
            headingToWrite = upper(strrep(sprintf(strrep(char(join(curFieldPath(1:(end-1)),'.')),'.',', ')),'_',' '));
    
            startCell = cellOffset(curOffset);
            xlswrite1(Excel,excelFileName,{headingToWrite},sheetNames{iFigure},startCell);
            Excel_utils.FormatCellFont(Excel, startCell, 'Calibri', 11, true, false, true, 'k');
    
            curOffset = curOffset + 2;%length(curFieldPath) +1;
    
            for iTable = 1:length(tableNames)
    
                fprintf('Currently writing table "%s"\n',tableNames{iTable});
    
                % write table here
                %                 headingToWrite
                %                 tableNames{iTable}
    
                xlswrite1(Excel,excelFileName,{upper(strrep(tableNames{iTable},'_',' '))},sheetNames{iFigure},sprintf('B%d',curOffset));
                Excel_utils.FormatCellFont(Excel, sprintf('B%d',curOffset), 'Calibri', 11, true, false, false, 'k');
                curOffset = curOffset+1;
                if isfield(curTables.(tableNames{iTable}),'colLabelHorizontal')
                    xlswrite1(Excel,excelFileName,{curTables.(tableNames{iTable}).colLabelHorizontal},sheetNames{iFigure},sprintf('C%d',curOffset));
                    Excel_utils.FormatCellFont(Excel, sprintf('C%d',curOffset), 'Calibri', 11, false, true, false, 'k');
                    curOffset = curOffset+1;
                end
    
                curTable = curTables.(tableNames{iTable}).table;
                tableToWrite = table2cell(curTable);
                startCell = cellOffset(curOffset);
    
                extraCol = 0;
                if ~isempty(curTable.Properties.RowNames)
                    extraCol = 1;
                end
                endCellCol = char(ExcelCol(1 + extraCol + size(curTable,2)));
    
                %                 xlswrite(excelFileName,tableToWrite,figureNames{iFigure},startCell);
                Excel.ActiveWorkbook.Save;
                Excel.ActiveWorkbook.Close;
                %                 invoke(Excel.Workbooks, 'Close');
                writetable(curTable,excelFileName,'Sheet',sheetNames{iFigure},'Range',startCell,'WriteRowNames',true);
                invoke(Excel.Workbooks, 'Open', excelFileNameAbs);
                endCell = [endCellCol,num2str(curOffset)];
                topRowRange = sprintf('%s:%s',startCell,endCell);
                Excel_utils.AlignCells(Excel, topRowRange, 3, false);
                Excel_utils.FormatCellFont(Excel, topRowRange, 'Calibri', 11, true, false, false, 'k');
    
                if ~isempty(curTable.Properties.RowNames)
                    xlswrite1(Excel,excelFileName,{''},sheetNames{iFigure},startCell); %delete the word "Row" that Matlab adds
                end
    
                endCellRow = curOffset + size(curTable,1);
                endCell = [endCellCol,num2str(endCellRow)];
                cellRange = sprintf('%s:%s',startCell,endCell);
                Excel_utils.FormatCellColor(Excel, cellRange, 15); % why 15? see https://analysistabs.com/excel-vba/colorindex/
                bottomCell = ['B',num2str(endCellRow)];
                leftColRange = sprintf('%s:%s',startCell,bottomCell);
    
                if ~isempty(curTable.Properties.RowNames)
                    Excel_utils.FormatCellFont(Excel, leftColRange, 'Calibri', 11, true, false, false, 'k');
                    Excel_utils.AlignCells(Excel, leftColRange, 4, false);
                end
    
                if isfield(curTables.(tableNames{iTable}),'rowLabelVertical')
                    xlswrite1(Excel,excelFileName,{curTables.(tableNames{iTable}).rowLabelVertical},sheetNames{iFigure},sprintf('A%d',curOffset+1));
                    Excel_utils.FormatCellFont(Excel, sprintf('A%d',curOffset+1), 'Calibri', 11, false, true, false, 'k');
                end
    
                curOffset = curOffset + size(curTable,1) + 2;
    
                legendToWrite = strrep(curTables.(tableNames{iTable}).legend,newline,' ');
    
                startCell = cellOffset(curOffset);
    
                xlswrite1(Excel,excelFileName,{legendToWrite},sheetNames{iFigure},startCell);
                Excel_utils.WrapText(Excel, startCell, true);
                Excel_utils.FormatCellFont(Excel, startCell, 'Calibri', 8, false, false, false, 'k');
                curOffset = curOffset + 2;%length(curFieldPath) +1;
            end
        end
        %         xlswrite(excelFileName,{''},1,'A1'); %move to first cell
    
        %          [tableData,pathToChild] = getNestedTableData(data.(figureNames{iFigure}),{});
        %          a = tableData
        %         for iTable = 1:length(tableData)
        % %            join(cellflat(pathToChild{iTable}),'.')
        %            tableName = tableData{iTable}
        %         end
        %         tableNames = fieldnames(data.(figureNames{iFigure}).tables);
        %         for iTable = 1:length(tableNames)
        %             xlswrite(excelFileName,table2cell(data.(figureNames{iFigure}).tables.(tableNames{iTable}).table),figureNames{iFigure});
        %         end
        Excel_utils.AutoSizeColumns(Excel);
        Excel_utils.AlignCells(Excel, '$A:$A', 4, false);
    end
    
    Excel.Range('A1').Select;
    Excel.Worksheets.Item(1).Activate;
    Excel.ActiveWorkbook.Save;
    
    Excel.Quit;
    delete(Excel);
    clear('Excel');
end

% StructFind()
% https://uk.mathworks.com/matlabcentral/fileexchange/35022-structfind?s_tid=FX_rc2_behav
% Accessed 27th September 2023
%
% Copyright (c) 2012, Alexander Mering
% All rights reserved.
%
% Redistribution and use in source and binary forms, with or without
% modification, are permitted provided that the following conditions are met:
%
% * Redistributions of source code must retain the above copyright notice, this
%   list of conditions and the following disclaimer.
%
% * Redistributions in binary form must reproduce the above copyright notice,
%   this list of conditions and the following disclaimer in the documentation
%   and/or other materials provided with the distribution
% THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
% AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
% IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
% DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE
% FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
% DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
% SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
% CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
% OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
% OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
%
function FoundFields = StructFind(search_struct,search_object,varargin)
% Function to search for entries within a structure. This is done
% recursively, running through all elements of array-structures.
%
% Input:
%
%       search_struct:      structure to be searched, could also be array of
%                           struct
%
%       search_object:      string, integer, cell, array or other thing to
%                           be searched for
%
%    [optional]
%
%       structure_name:     name of the structure to be searched. Used for
%                           full output of the structure content
%
% Output:
%
%       FoundFields:        Cell array of fields where the search object
%                           is found.
%
%
% Alexander Mering
% Version 1.0; 09. February 2012
%
    if ~isstruct(search_struct)
        error('Input should be a structure!')
        return
    end
    FoundFields = Really_StructFind(search_struct,search_object);
    % Prepend structure name to the output from the search function
    if nargin > 2 && length(varargin{1}) > 0
        FoundFields = strcat(repmat({varargin{1}},length(FoundFields),1),FoundFields');
    end
end
%% Search function
% used for recursive search through the structure
function FoundFieldsList = Really_StructFind(search_struct,search_field)
    % initialize output
    FoundFieldsList = cell('');
    % get fieldnames
    struct_fields  = fieldnames(search_struct);
    % outer loop is array of struct
    for n = 1:length(search_struct)
    
        % inner loop runs through the fields of the struct
        for m= 1: length(struct_fields)
    
            % get field to be worked with
            working_field = search_struct(n).(struct_fields{m});
    
            if isstruct(working_field) && ~isfield(working_field,search_field)
                % run search routine recursively
                InsideFound = Really_StructFind(working_field,search_field);
    
                % append search results to current output
                for k = 1:length(InsideFound)
                    %                 FoundFieldsList{end+1} = strcat('(',num2str(n),').',struct_fields{m},InsideFound{k});
                    FoundFieldsList{end+1} = strcat('.',struct_fields{m},InsideFound{k});
                end
    
            elseif isfield(working_field,search_field)%isequal(working_field,search_object)   % HIT!!
                % append found fields to the result cell
                %             FoundFieldsList{end+1} = strcat('(',num2str(n),').',struct_fields{m});
                FoundFieldsList{end+1} = strcat('.',struct_fields{m});
            end
    
        end
    end
end