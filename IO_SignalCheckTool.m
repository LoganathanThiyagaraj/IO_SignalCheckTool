  %*****************************************************
% Tool     : IO_SignalCheckTool
% Developer: Loganathan Thiyagaraj
% Owner    : Noopur Dosi
%*****************************************************

function varargout = IO_SignalCheckTool(varargin)
% IO_SIGNALCHECKTOOL MATLAB code for IO_SignalCheckTool.fig
%      IO_SIGNALCHECKTOOL, by itself, creates a new IO_SIGNALCHECKTOOL or raises the existing
%      singleton*.
%
%      H = IO_SIGNALCHECKTOOL returns the handle to a new IO_SIGNALCHECKTOOL or the handle to
%      the existing singleton*.
%
%      IO_SIGNALCHECKTOOL('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in IO_SIGNALCHECKTOOL.M with the given input arguments.
%
%      IO_SIGNALCHECKTOOL('Property','Value',...) creates a new IO_SIGNALCHECKTOOL or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before IO_SignalCheckTool_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to IO_SignalCheckTool_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help IO_SignalCheckTool

% Last Modified by GUIDE v2.5 21-Apr-2020 11:39:28

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @IO_SignalCheckTool_OpeningFcn, ...
                   'gui_OutputFcn',  @IO_SignalCheckTool_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before IO_SignalCheckTool is made visible.
function IO_SignalCheckTool_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to IO_SignalCheckTool (see VARARGIN)

% Choose default command line output for IO_SignalCheckTool
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes IO_SignalCheckTool wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = IO_SignalCheckTool_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;



function edit1_Callback(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit1 as text
%        str2double(get(hObject,'String')) returns contents of edit1 as a double



% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in BrowseCSVfile.
function BrowseCSVfile_Callback(hObject, eventdata, handles)
% hObject    handle to BrowseCSVfile (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global present SignalListfile selBasepath file1;
clc
selBasepath = uigetdir(path,'Select CSV file folder');%Saving CSV file folder path
addpath(genpath(selBasepath))%Adds CSV fiel folder path to the current folder
present=pwd;
[file1,path1]=uigetfile('*.csv','Select  SWC SignalListfile');%Selects CSV file
%file=strcat('\',file1);
file=strcat(path1,file1);
SignalListfile=file;
set(handles.edit1,'String',file1);%Displays CSV file name in gui
set(handles.BrowseCSVfile,'Enable','off')%Disable browse button


function edit2_Callback(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit2 as text
%        str2double(get(hObject,'String')) returns contents of edit2 as a double


% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in BrowseSignalDB.
function BrowseSignalDB_Callback(hObject, eventdata, handles)
% hObject    handle to BrowseSignalDB (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global SignalDBfile file2;
[file2,path1]=uigetfile('*.xlsx','Select SignalDB file');%Selects SignalDB file
file=strcat('\',file2);
file=strcat(path1,file);
SignalDBfile=file;
set(handles.edit2,'String',file2);%Displays DB file name in gui
set(handles.BrowseSignalDB,'Enable','off')%Disable browse button


% --- Executes on selection change in SWClistbox.
function SWClistbox_Callback(hObject, eventdata, handles)
% hObject    handle to SWClistbox (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns SWClistbox contents as cell array
%        contents{get(hObject,'Value')} returns selected item from SWClistbox
global SWC
swc=get(handles.SWClistbox,'String');%Gets the list of contents as array  in list box
idx=get(handles.SWClistbox,'Value');%Gets index of SWC component in list box
SWC=swc{idx};% Storing selected SWC in variable


% --- Executes during object creation, after setting all properties.
function SWClistbox_CreateFcn(hObject, eventdata, handles)
% hObject    handle to SWClistbox (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in partnumberlistbox.
function partnumberlistbox_Callback(hObject, eventdata, handles)
% hObject    handle to partnumberlistbox (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns partnumberlistbox contents as cell array
%        contents{get(hObject,'Value')} returns selected item from partnumberlistbox
global Partnumber
partnumber=get(handles.partnumberlistbox,'String');%Gets the list of contents as array  in list box
idx=get(handles.partnumberlistbox,'Value');%Gets index of part number in list box
Partnumber=partnumber{idx};% Storing selected part number in variable

% --- Executes during object creation, after setting all properties.
function partnumberlistbox_CreateFcn(hObject, eventdata, handles)
% hObject    handle to partnumberlistbox (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in LoadSWC_partnumber.
function LoadSWC_partnumber_Callback(hObject, eventdata, handles)
% hObject    handle to LoadSWC_partnumber (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global SignalDBfile Table  SignalListfile;
swclist={'OEM_DIAG','ActivationManagement','BRK_ENG','YAW','HMI','SEN_Setting','Temp_In','DiagControl','VehStatus_in','VehStatus_out'};
exl = actxserver('excel.application');
exlWkbk = exl.Workbooks;
exlFile = exlWkbk.Open(SignalDBfile);%Opens SignalDBfile
exlFile2 = exlWkbk.Open(SignalListfile);%Opens SignalListfile
exlFile.Activate
SignalDBsht=exlFile.Sheets.Item('SignalDB');%Selects SignalDB sheet
range1 = get(SignalDBsht,'Range','A:KL');%Selects range to extract
range1.AutoFilter%Removes Autofilter
range1.AutoFilter
clc
Table = readtable(SignalDBfile,'Sheet','SignalDB', 'ReadVariableNames', true);%Data from SignaDB sheet is stored in matlab table
exlFile2.Activate
%Table2 = readtable(SignalListfile,'ReadVariableNames', true);%Data from Signallistfile is stored in matlab table
exlFile.Close();
exlFile2.Close();
exl.Quit;
exl.delete;
idx3 = find(strcmp(Table.Properties.VariableNames,'L21BPRC'));%Extracts index of L21BPRC column
idx4 = find(strcmp(Table.Properties.VariableNames,'Reserve_'));%Extracts index of Reserve_ column 
a=1;
for i=idx3:idx4-1
    Vehlist{a}=Table.Properties.VariableNames{i};%Populates partnumber lists
    a=a+1;
end 
set(handles.SWClistbox,'String',swclist);%Displays SWC list in listbox
set(handles.partnumberlistbox,'String',Vehlist);%Displays partnumber list in listbox
set(handles.LoadSWC_partnumber,'Enable','off')%Disable Load SWC.P/N button

%set(handles.SWClistbox,'Value');

% --- Executes on button press in Execution.
function Execution_Callback(hObject, eventdata, handles)
% hObject    handle to Execution (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global Partnumber SWC  Table file1 file2
%assignin('base','Table2',Table2)%Stores Table2 in workspace
assignin('base','Table',Table)%Stores Tablein workspace
command1=strcat('Table.',SWC); 
command2=strcat('Table.',Partnumber); 
PortName=Table.PortName_C1A_;
AutosarSignalDataType=Table.AutosarSignalDataType;
Tx_Rx=evalin('base',command1);%Stores SWC column in array
Vehicle=evalin('base',command2);%Stores partnumber column in array
IO=Tx_Rx;
 for i=1:length(Tx_Rx)
   if strcmp(Tx_Rx{i},'R')
       IO{i} = 'Input';
   elseif strcmp(Tx_Rx{i},'T')
       IO{i} = 'Output';
   end
 end
DataType=AutosarSignalDataType;
DataType= regexprep(DataType,'B','boolean'); 
DataType= regexprep(DataType,'BUS','BUS');
for i=1:length(DataType)
    if strcmp(DataType{i},'booleanUS')
        DataType{i}=PortName{i};
    end
end
DataType= regexprep(DataType,'Fl','single');
DataType= regexprep(DataType,'FI32','FI32');
DataType= regexprep(DataType,'S16','int16');
DataType= regexprep(DataType,'S32','int32');
DataType= regexprep(DataType,'S8','int8');
DataType= regexprep(DataType,'U16','uint16');
DataType= regexprep(DataType,'U32','uint32');
DataType= regexprep(DataType,'U8','uint8');
a=1;
%Sorting required data
for i=1:length(Tx_Rx)
    if (strcmp(Tx_Rx{i},'R') || strcmp(Tx_Rx{i},'T')) && strcmp(Vehicle{i},'Y')%Filters 'T' ,'R' and 'Y' in swc and  partnumber column
    Templatedata{a,1}=PortName{i};
    Templatedata{a,2}=AutosarSignalDataType{i};
    Templatedata{a,3}=Tx_Rx{i};
    Templatedata{a,4}=IO{i};
    Templatedata{a,5}=DataType{i};
    a=a+1;
    end    
end
Templatedata=array2table(Templatedata,'VariableNames',{'Portname','AutosarSignalDataType','Tx_Rx','I_O','type'});

filename=strcat(SWC,'_SignalDB_Check.xlsx');
%writetable(Table2,filename,'Sheet','Template','Range','A1','WriteVariableNames',true);%Write data in signallistfile into report
writetable(Templatedata,filename,'Sheet','Template','Range','F1','WriteVariableNames',true);%write data in signaldb into report
clc

[dummy, dummy, raw] = xlsread(SignalListfile) ;%Reads ConstantListfile

table2=array2table(raw,'VariableNames',{'Model','Port','IO','Type'});%Stores data in SignalListfile in matlab table
filename1=strcat('\',filename);
filename=strcat(pwd,filename1);
xlswrite(filename,raw,'Template')%Write data  in ConstantListfile into report
assignin('base','outType',outType)
%Below code apply borders
    exl = actxserver('excel.application');
    exlWkbk = exl.Workbooks;
    exlFile = exlWkbk.Open(filename);
    exlFile.Activate

     exlFile.Sheets.Item('Template').Range('A1:D2000').Borders.Item('xlInsideHorizontal').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('A1:D2000').Borders.Item('xlInsideHorizontal').Weight = 2;
    exlFile.Sheets.Item('Template').Range('A1:D2000').Borders.Item('xlInsideVertical').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('A1:D2000').Borders.Item('xlInsideVertical').Weight = 2;
      exlFile.Sheets.Item('Template').Range('F1:J2000').Borders.Item('xlInsideHorizontal').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('F1:J2000').Borders.Item('xlInsideHorizontal').Weight = 2;
    exlFile.Sheets.Item('Template').Range('F1:J2000').Borders.Item('xlInsideVertical').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('F1:J2000').Borders.Item('xlInsideVertical').Weight = 2;
     exlFile.Sheets.Item('Template').Range('J1:J2000').Borders.Item('xlEdgeRight').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('J1:J2000').Borders.Item('xlEdgeRight').Weight = 2;
     exlFile.Sheets.Item('Template').Range('F1:F2000').Borders.Item('xlEdgeLeft').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('F1:F2000').Borders.Item('xlEdgeLeft').Weight = 2;
     exlFile.Sheets.Item('Template').Range('D1:D2000').Borders.Item('xlEdgeRight').LineStyle = 1;
     exlFile.Sheets.Item('Template').Range('D1:D2000').Borders.Item('xlEdgeRight').Weight = 2;
     exlFile.Sheets.Item('Template').Range('A1:D1').Interior.ColorIndex = 6;
     exlFile.Sheets.Item('Template').Range('F1:J1').Interior.ColorIndex = 6;
     
%      t=datetime('today');
%      day1=day(t,'name');
    
    exl.ActiveWindow.Zoom = 75;
    exlFile.Save();
    exlFile.Close();
    exl.Quit;
    exl.delete;
	%Below code writes information data in info sheet
    sheetname='info';
dat=datestr(now);
data1=cellstr(dat);
data2={'Date'};
xlswrite(char(filename),data1,sheetname,'B1');
xlswrite(char(filename),data2,sheetname,'A1');
t=datetime('today');
day1=day(t,'name');
data1=cellstr(day1);
data2={'Day'};
xlswrite(char(filename),data1,sheetname,'B2');
xlswrite(char(filename),data2,sheetname,'A2');
data1=cellstr(Partnumber);
data2={'Partnumber'};
xlswrite(char(filename),data1,sheetname,'B3');
xlswrite(char(filename),data2,sheetname,'A3');
data1=cellstr(SWC);
data2={'SWC'};
xlswrite(char(filename),data1,sheetname,'B4');
xlswrite(char(filename),data2,sheetname,'A4');
data1=cellstr(file1);
data2={'SignalList'};
xlswrite(char(filename),data1,sheetname,'B5');
xlswrite(char(filename),data2,sheetname,'A5');
data1=cellstr(file2);
data2={'dbfile'};
xlswrite(char(filename),data1,sheetname,'B6');
xlswrite(char(filename),data2,sheetname,'A6');
exl = actxserver('excel.application');
    exlWkbk = exl.Workbooks;
    exlFile = exlWkbk.Open(filename);
    exlFile.Activate
     exlFile.Sheets.Item('Sheet1').Delete;
     exlFile.Sheets.Item('Sheet2').Delete;
     exlFile.Sheets.Item('Sheet3').Delete;
     exlFile.Sheets.Item('Template').Activate
     %Portlist=Table2.Port;
     %Type=Table2.Type;
     Portlist=table2.Port;%Extracts partnumber 
Type=table2.Type;
     assignin('base','Portlist',Portlist)%Stores Tablein workspace
     assignin('base','Type',Type)%Stores Tablein workspace
     assignin('base','Templatedata',Templatedata)%Stores Tablein workspace
      for i=1:length(PortName)
          Templarr{i}=PortName{i};
      end
       assignin('base','Templatedata',Templatedata)%Stores Tablein workspace
     for i=1:height(Templatedata)
         if ~ismember(Templatedata{i,1},Portlist)  
     exlFile.Sheets.Item('Template').Range(sprintf('F%d',i+1)).Interior.ColorIndex = 3;
         end
         if ismember(Templatedata{i,1},Portlist)
              for j=2:length(Portlist)
                 if strcmp(Templatedata{i,1},Portlist(j)) && strcmp(Templatedata{i,5},Type(j))==0
                    exlFile.Sheets.Item('Template').Range(sprintf('D%d',j)).Interior.ColorIndex = 3; 
                 end
              end
         end
     end
      for i=2:length(Portlist)
          if ~ismember(Portlist{i},PortName)  
   exlFile.Sheets.Item('Template').Range(sprintf('B%d',i)).Interior.ColorIndex = 3;
   exlFile.Sheets.Item('Template').Range(sprintf('C%d',i)).Interior.ColorIndex = 3; 
   exlFile.Sheets.Item('Template').Range(sprintf('D%d',i)).Interior.ColorIndex = 3; 
         end
      end
     
  for j=2:length(Type)
    if strcmp(Type(j),'boolean') || strcmp(Type(j),'int16')|| strcmp(Type(j),'int32')|| strcmp(Type(j),'int8')|| strcmp(Type(j),'uint16')|| strcmp(Type(j),'uint32')|| strcmp(Type(j),'uint8')|| strcmp(Type(j),'single')
 continue;
    else
        exlFile.Sheets.Item('Template').Range(sprintf('D%d',j)).Interior.ColorIndex = 3; 
    end
 end          
     
exlFile.Save();
    exlFile.Close();
    exl.Quit;
    exl.delete;
    f = msgbox('Executed');
set(handles.Execution,'Enable','off')

% --- Executes on button press in End.
function End_Callback(hObject, eventdata, handles)
% hObject    handle to End (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
close(IO_SignalCheckTool)
