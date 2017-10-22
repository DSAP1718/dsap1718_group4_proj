%*************************************************************************%
% Title:        Speech Recorder for Speech Analysis
% Description:   Records a Speech and Exports the data to Excel tool which
%                will perform MFCC Analysis.
% Filename: speechAnalysis.m
% Version: v00.01
% Author:   Group 4
%           Alolor, Reynald
%           Camero, Jan Andrew
%           Catacutan, Jairus Roben
%           Ibe, John Edwin
%           Periabras, Redentor
%           Soria, Keira
% Yr&Sec: BSCS 4-4
% Subject: Digital Speech and Audio Signal Processing (DSAP)
%*************************************************************************%

function speechAnalysis(recordingLength)
%- Initialization --------------------------------------------------------%
    SAMPLING_RATE   = 32000; % 32 kHz     - samples per second
    BITS_PER_SAMPLE = 24;    % 24 bits    - for sampling accuracy
    CHANNEL         = 1;     % 1 for Mono
    
    %References:
    %   - http://wiki.audacityteam.org/wiki/Sample_Rates
    %   - http://www.resoundsound.com/sample-rate-bit-depth/
    
    EXCEL_TOOL = fullfile(pwd,'SpeechAnalysis_Group_4.xlsm');
%-------------------------------------------------------------------------%  
 
%- Recording -------------------------------------------------------------%
    recObj = audiorecorder(SAMPLING_RATE, BITS_PER_SAMPLE, CHANNEL);
    fprintf('* Recording');
    fprintf('\n\tStart speaking.');
    recordblocking(recObj, recordingLength);
    fprintf('\n\tStop Speaking.');
    
    fprintf('\n* Playback');
    play(recObj);
%-------------------------------------------------------------------------%

%- Exporting of Data and Simulation --------------------------------------%
    audioData = getaudiodata(recObj);
    
    %Check for existing Automation Server
    fprintf('\n\nConnecting to Excel Automation Server...');
    try
        %Use existing Automation Server
        excel = actxGetRunningServer('Excel.Application');
    catch
        %Open New Excel Automation Server
        excel = actxserver('Excel.Application');
    end
    
    try
        %Initialize setup for the ActiveXServer
        excel.DisplayAlerts = 0;    %Disable Alerts
        
        %Get all Workbooks
        workbooks = excel.Workbooks;
        
        %Make Excel Visible
        excel.Visible = 1;
        
        %Open Excel file
        fprintf('\nOpening to Excel Tool...');
        workbook = workbooks.Open(EXCEL_TOOL);
    
        %Specify sheet number and range
        sheetNumber = 1;
        range = strcat('A1:A',num2str(length(audioData)));

        %Make the first sheet active
        sheets = workbook.Sheets;
        sheet = get(sheets, 'Item', sheetNumber);
        invoke(sheet, 'Activate');
        activeSheet = workbook.Activesheet;

        %Export data from MATLAB to Excel Tool
        fprintf('\nExporting data to Excel Tool...');
        activeSheet_range = get(activeSheet, 'Range', range);
        set(activeSheet_range, 'Value', audioData);
        
        %Add Plot
        fprintf('\nCreating Plot...');
        sheet = workbook.Worksheets.Item('START >>');
        chart = sheet.ChartObjects.Item(1);
        chart.Chart.SetSourceData(activeSheet.Range(range));
        
        %Pause for exposure of Graph
        pause(10);
        
        %Switch to Sheet 2 for Button Controlled MFCC Generation
        fprintf('\nSwitching to Simulation Page...\n');
        sheet = get(sheets, 'Item', sheetNumber + 1);
        invoke(sheet, 'Activate');
    catch e
        %Close Excel Tool upon Encountering Error
        invoke(excel, 'Quit');
        clear Excel;
        fprintf('\nError:');
        fprintf('\t%s\n',e.message);
        disp('Stacktrace: ');
        disp(e.stack);
    end
%-------------------------------------------------------------------------%
end

