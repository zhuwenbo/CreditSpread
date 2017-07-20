function varargout = testorGUI(varargin)
% TESTORGUI MATLAB code for testorGUI.fig
%      TESTORGUI, by itself, creates a new TESTORGUI or raises the existing
%      singleton*.
%
%      H = TESTORGUI returns the handle to a new TESTORGUI or the handle to
%      the existing singleton*.
%
%      TESTORGUI('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in TESTORGUI.M with the given input arguments.
%
%      TESTORGUI('Property','Value',...) creates a new TESTORGUI or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before testorGUI_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to testorGUI_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help testorGUI

% Last Modified by GUIDE v2.5 07-Jul-2017 23:01:28

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @testorGUI_OpeningFcn, ...
                   'gui_OutputFcn',  @testorGUI_OutputFcn, ...
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


% --- Executes just before testorGUI is made visible.
function testorGUI_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to testorGUI (see VARARGIN)

% Choose default command line output for testorGUI
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes testorGUI wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = testorGUI_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;



function fedrate1_Callback(hObject, eventdata, handles)
% hObject    handle to fedrate1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of fedrate1 as text
%        str2double(get(hObject,'String')) returns contents of fedrate1 as a double


% --- Executes during object creation, after setting all properties.
function fedrate1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to fedrate1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function fedrate2_Callback(hObject, eventdata, handles)
% hObject    handle to fedrate2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of fedrate2 as text
%        str2double(get(hObject,'String')) returns contents of fedrate2 as a double


% --- Executes during object creation, after setting all properties.
function fedrate2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to fedrate2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function fedrate3_Callback(hObject, eventdata, handles)
% hObject    handle to fedrate3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of fedrate3 as text
%        str2double(get(hObject,'String')) returns contents of fedrate3 as a double


% --- Executes during object creation, after setting all properties.
function fedrate3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to fedrate3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function fedrate4_Callback(hObject, eventdata, handles)
% hObject    handle to fedrate4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of fedrate4 as text
%        str2double(get(hObject,'String')) returns contents of fedrate4 as a double


% --- Executes during object creation, after setting all properties.
function fedrate4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to fedrate4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function fedrate5_Callback(hObject, eventdata, handles)
% hObject    handle to fedrate5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of fedrate5 as text
%        str2double(get(hObject,'String')) returns contents of fedrate5 as a double


% --- Executes during object creation, after setting all properties.
function fedrate5_CreateFcn(hObject, eventdata, handles)
% hObject    handle to fedrate5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


function yr51_Callback(hObject, eventdata, handles)
% hObject    handle to yr51 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of yr51 as text
%        str2double(get(hObject,'String')) returns contents of yr51 as a double


% --- Executes during object creation, after setting all properties.
function yr51_CreateFcn(hObject, eventdata, handles)
% hObject    handle to yr51 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function yr101_Callback(hObject, eventdata, handles)
% hObject    handle to yr101 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of yr101 as text
%        str2double(get(hObject,'String')) returns contents of yr101 as a double


% --- Executes during object creation, after setting all properties.
function yr101_CreateFcn(hObject, eventdata, handles)
% hObject    handle to yr101 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


function md6mon_Callback(hObject, eventdata, handles)
% hObject    handle to md6mon (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of md6mon as text
%        str2double(get(hObject,'String')) returns contents of md6mon as a double


% --- Executes during object creation, after setting all properties.
function md6mon_CreateFcn(hObject, eventdata, handles)
% hObject    handle to md6mon (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function md1yr_Callback(hObject, eventdata, handles)
% hObject    handle to md1yr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of md1yr as text
%        str2double(get(hObject,'String')) returns contents of md1yr as a double


% --- Executes during object creation, after setting all properties.
function md1yr_CreateFcn(hObject, eventdata, handles)
% hObject    handle to md1yr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function md2yr_Callback(hObject, eventdata, handles)
% hObject    handle to md2yr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of md2yr as text
%        str2double(get(hObject,'String')) returns contents of md2yr as a double


% --- Executes during object creation, after setting all properties.
function md2yr_CreateFcn(hObject, eventdata, handles)
% hObject    handle to md2yr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function md3yr_Callback(hObject, eventdata, handles)
% hObject    handle to md3yr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of md3yr as text
%        str2double(get(hObject,'String')) returns contents of md3yr as a double


% --- Executes during object creation, after setting all properties.
function md3yr_CreateFcn(hObject, eventdata, handles)
% hObject    handle to md3yr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function md5yr_Callback(hObject, eventdata, handles)
% hObject    handle to md5yr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of md5yr as text
%        str2double(get(hObject,'String')) returns contents of md5yr as a double


% --- Executes during object creation, after setting all properties.
function md5yr_CreateFcn(hObject, eventdata, handles)
% hObject    handle to md5yr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function agencypct_Callback(hObject, eventdata, handles)
% hObject    handle to agencypct (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of agencypct as text
%        str2double(get(hObject,'String')) returns contents of agencypct as a double


% --- Executes during object creation, after setting all properties.
function agencypct_CreateFcn(hObject, eventdata, handles)
% hObject    handle to agencypct (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function extension_Callback(hObject, eventdata, handles)
% hObject    handle to extension (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of extension as text
%        str2double(get(hObject,'String')) returns contents of extension as a double


% --- Executes during object creation, after setting all properties.
function extension_CreateFcn(hObject, eventdata, handles)
% hObject    handle to extension (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in result.
function result_Callback(hObject, eventdata, handles)
% hObject    handle to result (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
[~, ~, raw, dates] = xlsread('C:\Users\Wenbo\Desktop\Amy_Project\regressionData.xlsx','Sheet1','A2:L380','',@convertSpreadsheetExcelDates);
raw = raw(:,[2,3,4,5,6,7,8,9,10,11,12]);
dates = dates(:,1);

data = reshape([raw{:}],size(raw));

Date = datetime([dates{:,1}].', 'ConvertFrom', 'Excel', 'Format', 'MM/dd/yyyy');
Mo1_r = data(:,1);
Mo3_r = data(:,2);
Mo6_r = data(:,3);
Yr1_r = data(:,4);
Yr2_r = data(:,5);
Yr3_r = data(:,6);
Yr5_r = data(:,7);
Yr7_r = data(:,8);
Yr10_r = data(:,9);
Yr20_r = data(:,10);
Yr30_r = data(:,11);

clearvars data raw dates;

set(handles.mon11,'string',Mo1_r(end));
set(handles.mon31,'string',Mo3_r(end));
set(handles.mon61,'string',Mo6_r(end));
set(handles.yr11,'string',Yr1_r(end));
set(handles.yr21,'string',Yr2_r(end));
set(handles.yr31,'string',Yr3_r(end));
set(handles.yr51,'string',Yr5_r(end));
set(handles.yr101,'string',Yr10_r(end));

date1 = Date(end) - days(1);
formatOut = 'mm/dd/yy';
set(handles.date1,'string',datestr(date1,formatOut));

date2 = dateshift(Date(end),'start','quarter','next');
date2 = date2 - days(1);
set(handles.date2,'string',datestr(date2,formatOut));

date3 = dateshift(date2 + days(1),'start','quarter','next');
date3 = date3 - days(1);
set(handles.date3,'string',datestr(date3,formatOut));

date4 = dateshift(date3 + days(1),'start','quarter','next');
date4 = date4 - days(1);
set(handles.date4,'string',datestr(date4,formatOut));

date5 = dateshift(date4+ days(1),'start','quarter','next');
date5 = date5 - days(1);
set(handles.date5,'string',datestr(date5,formatOut));


set(handles.cb1,'string',datestr(date1,formatOut));
set(handles.cb2,'string',datestr(date5,formatOut));

set(handles.aa1,'string',datestr(date1,formatOut));
set(handles.aa2,'string',datestr(date5,formatOut));

set(handles.ac1,'string',datestr(date1,formatOut));
set(handles.ac2,'string',datestr(date5,formatOut));

set(handles.nc1,'string',datestr(date1,formatOut));
set(handles.nc2,'string',datestr(date5,formatOut));


para = regress(Mo3_r,[Mo1_r,Yr10_r,ones(size(Mo1_r))]);
set(handles.mon35,'string',para(1)*1.13 + para(2)*3+para(3));

para = regress(Mo6_r,[Mo1_r,Yr10_r,ones(size(Mo1_r))]);
set(handles.mon65,'string', para(1)*1.13 + para(2)*3+para(3));

para = regress(Yr1_r,[Mo1_r,Yr10_r,ones(size(Mo1_r))]);
set(handles.yr15,'string', para(1)*1.13 + para(2)*3+para(3));

para = regress(Yr2_r,[Mo1_r,Yr10_r,ones(size(Mo1_r))]);
set(handles.yr25,'string', para(1)*1.13 + para(2)*3+para(3));

para = regress(Yr3_r,[Mo1_r,Yr10_r,ones(size(Mo1_r))]);
set(handles.yr35,'string', para(1)*1.13 + para(2)*3+para(3));

para = regress(Yr5_r,[Mo1_r,Yr10_r,ones(size(Mo1_r))]);
set(handles.yr55,'string', para(1)*1.13 + para(2)*3+para(3));

set(handles.mon12,'string',get(handles.fedrate2,'string'));
set(handles.mon13,'string',get(handles.fedrate3,'string'));
set(handles.mon14,'string',get(handles.fedrate4,'string'));
set(handles.mon15,'string',get(handles.fedrate5,'string'));

mon32 = (str2double(get(handles.mon35,'string'))-str2double(get(handles.mon31,'string')))...
    *(datenum(get(handles.date2,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.mon31,'string'));
set(handles.mon32,'string',mon32);

mon33 = (str2double(get(handles.mon35,'string'))-str2double(get(handles.mon31,'string')))...
    *(datenum(get(handles.date3,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.mon31,'string'));
set(handles.mon33,'string',mon33);

mon34 = (str2double(get(handles.mon35,'string'))-str2double(get(handles.mon31,'string')))...
    *(datenum(get(handles.date4,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.mon31,'string'));
set(handles.mon34,'string',mon34);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
mon62 = (str2double(get(handles.mon65,'string'))-str2double(get(handles.mon61,'string')))...
    *(datenum(get(handles.date2,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.mon61,'string'));
set(handles.mon62,'string',mon62);

mon63 = (str2double(get(handles.mon65,'string'))-str2double(get(handles.mon61,'string')))...
    *(datenum(get(handles.date3,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.mon61,'string'));
set(handles.mon63,'string',mon63);

mon64 = (str2double(get(handles.mon65,'string'))-str2double(get(handles.mon61,'string')))...
    *(datenum(get(handles.date4,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.mon61,'string'));
set(handles.mon64,'string',mon64);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
yr12 = (str2double(get(handles.yr15,'string'))-str2double(get(handles.yr11,'string')))...
    *(datenum(get(handles.date2,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr11,'string'));
set(handles.yr12,'string',yr12);

yr13 = (str2double(get(handles.yr15,'string'))-str2double(get(handles.yr11,'string')))...
    *(datenum(get(handles.date3,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr11,'string'));
set(handles.yr13,'string',yr13);

yr14 = (str2double(get(handles.yr15,'string'))-str2double(get(handles.yr11,'string')))...
    *(datenum(get(handles.date4,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr11,'string'));
set(handles.yr14,'string',yr14);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
yr22 = (str2double(get(handles.yr25,'string'))-str2double(get(handles.yr21,'string')))...
    *(datenum(get(handles.date2,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr21,'string'));
set(handles.yr22,'string',yr22);

yr23 = (str2double(get(handles.yr25,'string'))-str2double(get(handles.yr21,'string')))...
    *(datenum(get(handles.date3,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr21,'string'));
set(handles.yr23,'string',yr23);

yr24 = (str2double(get(handles.yr25,'string'))-str2double(get(handles.yr21,'string')))...
    *(datenum(get(handles.date4,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr21,'string'));
set(handles.yr24,'string',yr24);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
yr32 = (str2double(get(handles.yr35,'string'))-str2double(get(handles.yr31,'string')))...
    *(datenum(get(handles.date2,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr31,'string'));
set(handles.yr32,'string',yr32);

yr33 = (str2double(get(handles.yr35,'string'))-str2double(get(handles.yr31,'string')))...
    *(datenum(get(handles.date3,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr31,'string'));
set(handles.yr33,'string',yr33);

yr34 = (str2double(get(handles.yr35,'string'))-str2double(get(handles.yr31,'string')))...
    *(datenum(get(handles.date4,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr31,'string'));
set(handles.yr34,'string',yr34);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
yr52 = (str2double(get(handles.yr55,'string'))-str2double(get(handles.yr51,'string')))...
    *(datenum(get(handles.date2,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr51,'string'));
set(handles.yr52,'string',yr52);

yr53 = (str2double(get(handles.yr55,'string'))-str2double(get(handles.yr51,'string')))...
    *(datenum(get(handles.date3,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr51,'string'));
set(handles.yr53,'string',yr53);

yr54 = (str2double(get(handles.yr55,'string'))-str2double(get(handles.yr51,'string')))...
    *(datenum(get(handles.date4,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr51,'string'));
set(handles.yr54,'string',yr54);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
yr102 = (str2double(get(handles.yr105,'string'))-str2double(get(handles.yr101,'string')))...
    *(datenum(get(handles.date2,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr101,'string'));
set(handles.yr102,'string',yr102);

yr103 = (str2double(get(handles.yr105,'string'))-str2double(get(handles.yr101,'string')))...
    *(datenum(get(handles.date3,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr101,'string'));
set(handles.yr103,'string',yr103);

yr104 = (str2double(get(handles.yr105,'string'))-str2double(get(handles.yr101,'string')))...
    *(datenum(get(handles.date4,'string'))-datenum(get(handles.date1,'string')))...
    /(datenum(get(handles.date5,'string'))-datenum(get(handles.date1,'string')))...
    + str2double(get(handles.yr101,'string'));
set(handles.yr104,'string',yr104);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%







% Mo3(i) = (Mo3(5)-Mo3(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Mo3(1);
% Mo6(i) = (Mo6(5)-Mo6(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Mo6(1);
% Yr1(i) = (Yr1(5)-Yr1(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Yr1(1);
% Yr2(i) = (Yr2(5)-Yr2(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Yr2(1);
% Yr3(i) = (Yr3(5)-Yr3(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Yr3(1);
% Yr5(i) = (Yr5(5)-Yr5(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Yr5(1);
% Yr10(i) = (Yr10(5)-Yr10(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Yr10(1);


% --- Executes on button press in filltable.
function filltable_Callback(hObject, eventdata, handles)
% hObject    handle to filltable (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
data = get(handles.bigtable,'data');
for i = 1:7
    for j = [5,8,11]
        data((j-2):(j-1),i) = {num2str((str2double(data{j,i})-str2double(data{j-3,i}))/3*1 + str2double(data{(j-3),i}))...
            ;num2str((str2double(data{j,i})-str2double(data{j-3,i}))/3*2 + str2double(data{(j-3),i}))};
    end
end

set(handles.bigtable,'data',data);


% --- Executes on button press in interesrateimpact.
function interesrateimpact_Callback(hObject, eventdata, handles)
% hObject    handle to interesrateimpact (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%data = get(handles.bigtable,'data');

data(1,:) = [str2double(get(handles.mon11,'string'));str2double(get(handles.mon31,'string'));str2double(get(handles.mon61,'string'))...
    ;str2double(get(handles.yr11,'string'));str2double(get(handles.yr21,'string'));str2double(get(handles.yr31,'string'))...
    ;str2double(get(handles.yr51,'string'))];
data(2,:) = [str2double(get(handles.mon12,'string'));str2double(get(handles.mon32,'string'));str2double(get(handles.mon62,'string'))...
    ;str2double(get(handles.yr12,'string'));str2double(get(handles.yr22,'string'));str2double(get(handles.yr32,'string'))...
    ;str2double(get(handles.yr52,'string'))];
data(5,:) = [str2double(get(handles.mon13,'string'));str2double(get(handles.mon33,'string'));str2double(get(handles.mon63,'string'))...
    ;str2double(get(handles.yr13,'string'));str2double(get(handles.yr23,'string'));str2double(get(handles.yr33,'string'))...
    ;str2double(get(handles.yr53,'string'))];
data(8,:) = [str2double(get(handles.mon14,'string'));str2double(get(handles.mon34,'string'));str2double(get(handles.mon64,'string'))...
    ;str2double(get(handles.yr14,'string'));str2double(get(handles.yr24,'string'));str2double(get(handles.yr34,'string'))...
    ;str2double(get(handles.yr54,'string'))];
data(11,:) = [str2double(get(handles.mon15,'string'));str2double(get(handles.mon35,'string'));str2double(get(handles.mon65,'string'))...
    ;str2double(get(handles.yr15,'string'));str2double(get(handles.yr25,'string'));str2double(get(handles.yr35,'string'))...
    ;str2double(get(handles.yr55,'string'))];

for i = 1:7
    for j = [5,8,11]
        data((j-2):(j-1),i) = [((data(j,i))-(data(j-3,i)))/3*1 + data((j-3),i)...
            ;(data(j,i)-data(j-3,i))/3*2 + data((j-3),i)];
    end
end


% data = get(handles.bigtable,'data');
intdata = get(handles.inttable,'data');

income_col_6mo_1 = 1 + data(1,3)/400*1;
income_col_6mo_2 = income_col_6mo_1*0.75*(1+ data(1,3)/400)+ income_col_6mo_1*0.25*(1+data(4,3)/400);
income_col_6mo_3 = income_col_6mo_2*0.5*(1 + data(1,3)/400)+ income_col_6mo_2*0.25*(1+data(4,3)/400) + income_col_6mo_2*0.25*(1+data(7,3)/400);
income_col_6mo_final = income_col_6mo_3*0.25*(1 + data(1,3)/400)+ income_col_6mo_3*0.25*(1+data(4,3)/400) + income_col_6mo_3*0.25*(1+data(7,3)/400)...
    + income_col_6mo_3*0.25*(1+data(10,3)/400);
intdata(1,1) = {num2str(income_col_6mo_final-1)};

income_col_1yr_1 = 1 + data(1,4)/400*1;
income_col_1yr_2 = income_col_1yr_1*0.8*(1+ data(1,4)/400)+ income_col_1yr_1*0.2*(1+data(4,4)/400);
income_col_1yr_3 = income_col_1yr_2*0.6*(1 + data(1,4)/400)+ income_col_1yr_2*0.2*(1+data(4,4)/400) + income_col_1yr_2*0.2*(1+data(7,4)/400);
income_col_1yr_final = income_col_1yr_3*0.4*(1 + data(1,4)/400)+ income_col_1yr_3*0.2*(1+data(4,4)/400) + income_col_1yr_3*0.2*(1+data(7,4)/400)...
    + income_col_1yr_3*0.2*(1+data(10,4)/400);
intdata(1,2) = {num2str(income_col_1yr_final-1)};

income_col_2yr_1 = 1 + data(1,5)/400*1;
income_col_2yr_2 = income_col_2yr_1*0.875*(1+ data(1,5)/400)+ income_col_2yr_1*0.125*(1+data(4,5)/400);
income_col_2yr_3 = income_col_2yr_2*0.75*(1 + data(1,5)/400)+ income_col_2yr_2*0.125*(1+data(4,5)/400) + income_col_2yr_2*0.125*(1+data(7,5)/400);
income_col_2yr_final = income_col_2yr_3*0.625*(1 + data(1,5)/400)+ income_col_2yr_3*0.125*(1+data(4,5)/400) + income_col_2yr_3*0.125*(1+data(7,5)/400)...
    + income_col_2yr_3*0.125*(1+data(10,5)/400);
intdata(1,3) = {num2str(income_col_2yr_final-1)};

income_col_3yr_1 = 1 + data(1,6)/400*1;
income_col_3yr_2 = income_col_3yr_1*0.92*(1+ data(1,6)/400)+ income_col_3yr_1*0.08*(1+data(4,6)/400);
income_col_3yr_3 = income_col_3yr_2*0.84*(1 + data(1,6)/400)+ income_col_3yr_2*0.08*(1+data(4,6)/400) + income_col_3yr_2*0.08*(1+data(7,6)/400);
income_col_3yr_final = income_col_3yr_3*0.76*(1 + data(1,6)/400)+ income_col_3yr_3*0.08*(1+data(4,6)/400) + income_col_3yr_3*0.08*(1+data(7,6)/400)...
    + income_col_3yr_3*0.08*(1+data(10,6)/400);
intdata(1,4) = {num2str(income_col_3yr_final-1)};

income_col_5yr_1 = 1 + data(1,7)/400*1;
income_col_5yr_2 = income_col_5yr_1*0.95*(1+ data(1,7)/400)+ income_col_5yr_1*0.05*(1+data(4,7)/400);
income_col_5yr_3 = income_col_5yr_2*0.9*(1 + data(1,7)/400)+ income_col_5yr_2*0.05*(1+data(4,7)/400) + income_col_5yr_2*0.05*(1+data(7,7)/400);
income_col_5yr_final = income_col_5yr_3*0.85*(1 + data(1,7)/400)+ income_col_5yr_3*0.05*(1+data(4,7)/400) + income_col_5yr_3*0.05*(1+data(7,7)/400)...
    + income_col_5yr_3*0.05*(1+data(10,7)/400);
intdata(1,5) = {num2str(income_col_5yr_final-1)};

Agency_RMBS_pct = str2double(get(handles.agencypct,'string'));
Agency_RMBS_DurationExtension = str2double(get(handles.extension,'string'));

intdata(2,1) = {num2str((str2double(get(handles.mon65,'string'))-(str2double(get(handles.mon61,'string'))))*str2double(get(handles.md6mon,'string'))/-100*...
    (1-Agency_RMBS_pct)+(str2double(get(handles.mon65,'string'))-(str2double(get(handles.mon61,'string'))))*str2double(get(handles.md6mon,'string'))/-100*Agency_RMBS_pct*Agency_RMBS_DurationExtension)};

intdata(2,2) = {num2str((str2double(get(handles.yr15,'string'))-(str2double(get(handles.yr11,'string'))))*str2double(get(handles.md1yr,'string'))/-100*...
    (1-Agency_RMBS_pct)+(str2double(get(handles.yr15,'string'))-(str2double(get(handles.yr11,'string'))))*str2double(get(handles.md1yr,'string'))/-100*Agency_RMBS_pct*Agency_RMBS_DurationExtension)};

intdata(2,3) = {num2str((str2double(get(handles.yr25,'string'))-(str2double(get(handles.yr21,'string'))))*str2double(get(handles.md2yr,'string'))/-100*...
    (1-Agency_RMBS_pct)+(str2double(get(handles.yr25,'string'))-(str2double(get(handles.yr21,'string'))))*str2double(get(handles.md2yr,'string'))/-100*Agency_RMBS_pct*Agency_RMBS_DurationExtension)};

intdata(2,4) = {num2str((str2double(get(handles.yr35,'string'))-(str2double(get(handles.yr31,'string'))))*str2double(get(handles.md3yr,'string'))/-100*...
    (1-Agency_RMBS_pct)+(str2double(get(handles.yr35,'string'))-(str2double(get(handles.yr31,'string'))))*str2double(get(handles.md3yr,'string'))/-100*Agency_RMBS_pct*Agency_RMBS_DurationExtension)};

intdata(2,5) = {num2str((str2double(get(handles.yr55,'string'))-(str2double(get(handles.yr51,'string'))))*str2double(get(handles.md5yr,'string'))/-100*...
    (1-Agency_RMBS_pct)+(str2double(get(handles.yr55,'string'))-(str2double(get(handles.yr51,'string'))))*str2double(get(handles.md5yr,'string'))/-100*Agency_RMBS_pct*Agency_RMBS_DurationExtension)};


intdata(3,1) = {num2str(str2double(intdata(2,1)) + str2double(intdata(1,1)))};
intdata(3,2) = {num2str(str2double(intdata(2,2)) + str2double(intdata(1,2)))};
intdata(3,3) = {num2str(str2double(intdata(2,3)) + str2double(intdata(1,3)))};
intdata(3,4) = {num2str(str2double(intdata(2,4)) + str2double(intdata(1,4)))};
intdata(3,5) = {num2str(str2double(intdata(2,5)) + str2double(intdata(1,5)))};



set(handles.inttable,'data',intdata);



function corpratepct_Callback(hObject, eventdata, handles)
% hObject    handle to corpratepct (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of corpratepct as text
%        str2double(get(hObject,'String')) returns contents of corpratepct as a double


% --- Executes during object creation, after setting all properties.
function corpratepct_CreateFcn(hObject, eventdata, handles)
% hObject    handle to corpratepct (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function abspct_Callback(hObject, eventdata, handles)
% hObject    handle to abspct (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of abspct as text
%        str2double(get(hObject,'String')) returns contents of abspct as a double


% --- Executes during object creation, after setting all properties.
function abspct_CreateFcn(hObject, eventdata, handles)
% hObject    handle to abspct (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function cmbspct_Callback(hObject, eventdata, handles)
% hObject    handle to cmbspct (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of cmbspct as text
%        str2double(get(hObject,'String')) returns contents of cmbspct as a double


% --- Executes during object creation, after setting all properties.
function cmbspct_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cmbspct (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function corp6_Callback(hObject, eventdata, handles)
% hObject    handle to corp6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of corp6 as text
%        str2double(get(hObject,'String')) returns contents of corp6 as a double


% --- Executes during object creation, after setting all properties.
function corp6_CreateFcn(hObject, eventdata, handles)
% hObject    handle to corp6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function corp62_Callback(hObject, eventdata, handles)
% hObject    handle to corp62 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of corp62 as text
%        str2double(get(hObject,'String')) returns contents of corp62 as a double


% --- Executes during object creation, after setting all properties.
function corp62_CreateFcn(hObject, eventdata, handles)
% hObject    handle to corp62 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function abs6_Callback(hObject, eventdata, handles)
% hObject    handle to abs6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of abs6 as text
%        str2double(get(hObject,'String')) returns contents of abs6 as a double


% --- Executes during object creation, after setting all properties.
function abs6_CreateFcn(hObject, eventdata, handles)
% hObject    handle to abs6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function abs62_Callback(hObject, eventdata, handles)
% hObject    handle to abs62 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of abs62 as text
%        str2double(get(hObject,'String')) returns contents of abs62 as a double


% --- Executes during object creation, after setting all properties.
function abs62_CreateFcn(hObject, eventdata, handles)
% hObject    handle to abs62 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function agency6_Callback(hObject, eventdata, handles)
% hObject    handle to agency6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of agency6 as text
%        str2double(get(hObject,'String')) returns contents of agency6 as a double


% --- Executes during object creation, after setting all properties.
function agency6_CreateFcn(hObject, eventdata, handles)
% hObject    handle to agency6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function agency62_Callback(hObject, eventdata, handles)
% hObject    handle to agency62 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of agency62 as text
%        str2double(get(hObject,'String')) returns contents of agency62 as a double


% --- Executes during object creation, after setting all properties.
function agency62_CreateFcn(hObject, eventdata, handles)
% hObject    handle to agency62 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function cmbs6_Callback(hObject, eventdata, handles)
% hObject    handle to cmbs6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of cmbs6 as text
%        str2double(get(hObject,'String')) returns contents of cmbs6 as a double


% --- Executes during object creation, after setting all properties.
function cmbs6_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cmbs6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function cmbs62_Callback(hObject, eventdata, handles)
% hObject    handle to cmbs62 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of cmbs62 as text
%        str2double(get(hObject,'String')) returns contents of cmbs62 as a double


% --- Executes during object creation, after setting all properties.
function cmbs62_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cmbs62 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function corp1_Callback(hObject, eventdata, handles)
% hObject    handle to corp1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of corp1 as text
%        str2double(get(hObject,'String')) returns contents of corp1 as a double


% --- Executes during object creation, after setting all properties.
function corp1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to corp1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function corp12_Callback(hObject, eventdata, handles)
% hObject    handle to corp12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of corp12 as text
%        str2double(get(hObject,'String')) returns contents of corp12 as a double


% --- Executes during object creation, after setting all properties.
function corp12_CreateFcn(hObject, eventdata, handles)
% hObject    handle to corp12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function abs1_Callback(hObject, eventdata, handles)
% hObject    handle to abs1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of abs1 as text
%        str2double(get(hObject,'String')) returns contents of abs1 as a double


% --- Executes during object creation, after setting all properties.
function abs1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to abs1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function abs12_Callback(hObject, eventdata, handles)
% hObject    handle to abs12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of abs12 as text
%        str2double(get(hObject,'String')) returns contents of abs12 as a double


% --- Executes during object creation, after setting all properties.
function abs12_CreateFcn(hObject, eventdata, handles)
% hObject    handle to abs12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function agency1_Callback(hObject, eventdata, handles)
% hObject    handle to agency1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of agency1 as text
%        str2double(get(hObject,'String')) returns contents of agency1 as a double


% --- Executes during object creation, after setting all properties.
function agency1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to agency1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function agency12_Callback(hObject, eventdata, handles)
% hObject    handle to agency12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of agency12 as text
%        str2double(get(hObject,'String')) returns contents of agency12 as a double


% --- Executes during object creation, after setting all properties.
function agency12_CreateFcn(hObject, eventdata, handles)
% hObject    handle to agency12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function cmbs1_Callback(hObject, eventdata, handles)
% hObject    handle to cmbs1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of cmbs1 as text
%        str2double(get(hObject,'String')) returns contents of cmbs1 as a double


% --- Executes during object creation, after setting all properties.
function cmbs1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cmbs1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function cmbs12_Callback(hObject, eventdata, handles)
% hObject    handle to cmbs12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of cmbs12 as text
%        str2double(get(hObject,'String')) returns contents of cmbs12 as a double


% --- Executes during object creation, after setting all properties.
function cmbs12_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cmbs12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function corp2_Callback(hObject, eventdata, handles)
% hObject    handle to corp2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of corp2 as text
%        str2double(get(hObject,'String')) returns contents of corp2 as a double


% --- Executes during object creation, after setting all properties.
function corp2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to corp2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function corp22_Callback(hObject, eventdata, handles)
% hObject    handle to corp22 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of corp22 as text
%        str2double(get(hObject,'String')) returns contents of corp22 as a double


% --- Executes during object creation, after setting all properties.
function corp22_CreateFcn(hObject, eventdata, handles)
% hObject    handle to corp22 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function abs2_Callback(hObject, eventdata, handles)
% hObject    handle to abs2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of abs2 as text
%        str2double(get(hObject,'String')) returns contents of abs2 as a double


% --- Executes during object creation, after setting all properties.
function abs2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to abs2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function abs22_Callback(hObject, eventdata, handles)
% hObject    handle to abs22 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of abs22 as text
%        str2double(get(hObject,'String')) returns contents of abs22 as a double


% --- Executes during object creation, after setting all properties.
function abs22_CreateFcn(hObject, eventdata, handles)
% hObject    handle to abs22 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function agency2_Callback(hObject, eventdata, handles)
% hObject    handle to agency2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of agency2 as text
%        str2double(get(hObject,'String')) returns contents of agency2 as a double


% --- Executes during object creation, after setting all properties.
function agency2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to agency2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function agency22_Callback(hObject, eventdata, handles)
% hObject    handle to agency22 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of agency22 as text
%        str2double(get(hObject,'String')) returns contents of agency22 as a double


% --- Executes during object creation, after setting all properties.
function agency22_CreateFcn(hObject, eventdata, handles)
% hObject    handle to agency22 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function cmbs2_Callback(hObject, eventdata, handles)
% hObject    handle to cmbs2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of cmbs2 as text
%        str2double(get(hObject,'String')) returns contents of cmbs2 as a double


% --- Executes during object creation, after setting all properties.
function cmbs2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cmbs2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function cmbs22_Callback(hObject, eventdata, handles)
% hObject    handle to cmbs22 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of cmbs22 as text
%        str2double(get(hObject,'String')) returns contents of cmbs22 as a double


% --- Executes during object creation, after setting all properties.
function cmbs22_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cmbs22 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function corp3_Callback(hObject, eventdata, handles)
% hObject    handle to corp3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of corp3 as text
%        str2double(get(hObject,'String')) returns contents of corp3 as a double


% --- Executes during object creation, after setting all properties.
function corp3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to corp3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function corp32_Callback(hObject, eventdata, handles)
% hObject    handle to corp32 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of corp32 as text
%        str2double(get(hObject,'String')) returns contents of corp32 as a double


% --- Executes during object creation, after setting all properties.
function corp32_CreateFcn(hObject, eventdata, handles)
% hObject    handle to corp32 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function abs3_Callback(hObject, eventdata, handles)
% hObject    handle to abs3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of abs3 as text
%        str2double(get(hObject,'String')) returns contents of abs3 as a double


% --- Executes during object creation, after setting all properties.
function abs3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to abs3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function abs32_Callback(hObject, eventdata, handles)
% hObject    handle to abs32 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of abs32 as text
%        str2double(get(hObject,'String')) returns contents of abs32 as a double


% --- Executes during object creation, after setting all properties.
function abs32_CreateFcn(hObject, eventdata, handles)
% hObject    handle to abs32 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function agency3_Callback(hObject, eventdata, handles)
% hObject    handle to agency3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of agency3 as text
%        str2double(get(hObject,'String')) returns contents of agency3 as a double


% --- Executes during object creation, after setting all properties.
function agency3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to agency3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function agency32_Callback(hObject, eventdata, handles)
% hObject    handle to agency32 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of agency32 as text
%        str2double(get(hObject,'String')) returns contents of agency32 as a double


% --- Executes during object creation, after setting all properties.
function agency32_CreateFcn(hObject, eventdata, handles)
% hObject    handle to agency32 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function cmbs3_Callback(hObject, eventdata, handles)
% hObject    handle to cmbs3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of cmbs3 as text
%        str2double(get(hObject,'String')) returns contents of cmbs3 as a double


% --- Executes during object creation, after setting all properties.
function cmbs3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cmbs3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function cmbs32_Callback(hObject, eventdata, handles)
% hObject    handle to cmbs32 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of cmbs32 as text
%        str2double(get(hObject,'String')) returns contents of cmbs32 as a double


% --- Executes during object creation, after setting all properties.
function cmbs32_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cmbs32 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function corp5_Callback(hObject, eventdata, handles)
% hObject    handle to corp5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of corp5 as text
%        str2double(get(hObject,'String')) returns contents of corp5 as a double


% --- Executes during object creation, after setting all properties.
function corp5_CreateFcn(hObject, eventdata, handles)
% hObject    handle to corp5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function corp52_Callback(hObject, eventdata, handles)
% hObject    handle to corp52 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of corp52 as text
%        str2double(get(hObject,'String')) returns contents of corp52 as a double


% --- Executes during object creation, after setting all properties.
function corp52_CreateFcn(hObject, eventdata, handles)
% hObject    handle to corp52 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function abs5_Callback(hObject, eventdata, handles)
% hObject    handle to abs5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of abs5 as text
%        str2double(get(hObject,'String')) returns contents of abs5 as a double


% --- Executes during object creation, after setting all properties.
function abs5_CreateFcn(hObject, eventdata, handles)
% hObject    handle to abs5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function abs52_Callback(hObject, eventdata, handles)
% hObject    handle to abs52 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of abs52 as text
%        str2double(get(hObject,'String')) returns contents of abs52 as a double


% --- Executes during object creation, after setting all properties.
function abs52_CreateFcn(hObject, eventdata, handles)
% hObject    handle to abs52 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function agency5_Callback(hObject, eventdata, handles)
% hObject    handle to agency5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of agency5 as text
%        str2double(get(hObject,'String')) returns contents of agency5 as a double


% --- Executes during object creation, after setting all properties.
function agency5_CreateFcn(hObject, eventdata, handles)
% hObject    handle to agency5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function agency52_Callback(hObject, eventdata, handles)
% hObject    handle to agency52 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of agency52 as text
%        str2double(get(hObject,'String')) returns contents of agency52 as a double


% --- Executes during object creation, after setting all properties.
function agency52_CreateFcn(hObject, eventdata, handles)
% hObject    handle to agency52 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function cmbs5_Callback(hObject, eventdata, handles)
% hObject    handle to cmbs5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of cmbs5 as text
%        str2double(get(hObject,'String')) returns contents of cmbs5 as a double


% --- Executes during object creation, after setting all properties.
function cmbs5_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cmbs5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function cmbs52_Callback(hObject, eventdata, handles)
% hObject    handle to cmbs52 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of cmbs52 as text
%        str2double(get(hObject,'String')) returns contents of cmbs52 as a double


% --- Executes during object creation, after setting all properties.
function cmbs52_CreateFcn(hObject, eventdata, handles)
% hObject    handle to cmbs52 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function agencyspread_Callback(hObject, eventdata, handles)
% hObject    handle to agencyspread (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of agencyspread as text
%        str2double(get(hObject,'String')) returns contents of agencyspread as a double


% --- Executes during object creation, after setting all properties.
function agencyspread_CreateFcn(hObject, eventdata, handles)
% hObject    handle to agencyspread (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function sp6_Callback(hObject, eventdata, handles)
% hObject    handle to sp6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of sp6 as text
%        str2double(get(hObject,'String')) returns contents of sp6 as a double


% --- Executes during object creation, after setting all properties.
function sp6_CreateFcn(hObject, eventdata, handles)
% hObject    handle to sp6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function sp1_Callback(hObject, eventdata, handles)
% hObject    handle to sp1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of sp1 as text
%        str2double(get(hObject,'String')) returns contents of sp1 as a double


% --- Executes during object creation, after setting all properties.
function sp1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to sp1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function sp2_Callback(hObject, eventdata, handles)
% hObject    handle to sp2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of sp2 as text
%        str2double(get(hObject,'String')) returns contents of sp2 as a double


% --- Executes during object creation, after setting all properties.
function sp2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to sp2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function sp3_Callback(hObject, eventdata, handles)
% hObject    handle to sp3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of sp3 as text
%        str2double(get(hObject,'String')) returns contents of sp3 as a double


% --- Executes during object creation, after setting all properties.
function sp3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to sp3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function sp5_Callback(hObject, eventdata, handles)
% hObject    handle to sp5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of sp5 as text
%        str2double(get(hObject,'String')) returns contents of sp5 as a double


% --- Executes during object creation, after setting all properties.
function sp5_CreateFcn(hObject, eventdata, handles)
% hObject    handle to sp5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in credit.
function credit_Callback(hObject, eventdata, handles)
% hObject    handle to credit (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
creditdata = get(handles.credittable,'data');
intdata = get(handles.inttable,'data');

Corp_CS_6Mon = str2double(get(handles.corp6,'string')) + str2double(get(handles.corp62,'string'));
ABS_CS_6Mon = str2double(get(handles.abs6,'string')) + str2double(get(handles.abs62,'string'));
MBS_CS_6Mon = str2double(get(handles.agency6,'string')) + str2double(get(handles.agency62,'string'));
CMBS_CS_6Mon = str2double(get(handles.cmbs6,'string')) + str2double(get(handles.cmbs62,'string'));

Corp_CS_1Yr = str2double(get(handles.corp1,'string')) + str2double(get(handles.corp12,'string'));
ABS_CS_1Yr = str2double(get(handles.abs1,'string')) + str2double(get(handles.abs12,'string'));
MBS_CS_1Yr = str2double(get(handles.agency1,'string')) + str2double(get(handles.agency12,'string'));
CMBS_CS_1Yr = str2double(get(handles.cmbs1,'string')) + str2double(get(handles.cmbs12,'string'));

Corp_CS_2Yr = str2double(get(handles.corp2,'string')) + str2double(get(handles.corp22,'string'));
ABS_CS_2Yr = str2double(get(handles.abs2,'string')) + str2double(get(handles.abs22,'string'));
MBS_CS_2Yr = str2double(get(handles.agency2,'string')) + str2double(get(handles.agency22,'string'));
CMBS_CS_2Yr = str2double(get(handles.cmbs2,'string')) + str2double(get(handles.cmbs22,'string'));

Corp_CS_3Yr = str2double(get(handles.corp3,'string')) + str2double(get(handles.corp32,'string'));
ABS_CS_3Yr = str2double(get(handles.abs3,'string')) + str2double(get(handles.abs32,'string'));
MBS_CS_3Yr = str2double(get(handles.agency3,'string')) + str2double(get(handles.agency32,'string'));
CMBS_CS_3Yr = str2double(get(handles.cmbs3,'string')) + str2double(get(handles.cmbs32,'string'));

Corp_CS_5Yr = str2double(get(handles.corp5,'string')) + str2double(get(handles.corp52,'string'));
ABS_CS_5Yr = str2double(get(handles.abs5,'string')) + str2double(get(handles.abs52,'string'));
MBS_CS_5Yr = str2double(get(handles.agency5,'string')) + str2double(get(handles.agency52,'string'));
CMBS_CS_5Yr = str2double(get(handles.cmbs5,'string')) + str2double(get(handles.cmbs52,'string'));

CorpBond_CommercialPaper_pct_sp = str2double(get(handles.corpratepct,'string'));
ABS_pct_sp = str2double(get(handles.abspct,'string'));
AgencyRMBS_pct_sp = str2double(get(handles.agencyspread,'string'));
CMBS_pct_sp = str2double(get(handles.cmbspct,'string'));

creditdata(1,1) = {num2str((1+sum(Corp_CS_6Mon)/2/100).^0.9*CorpBond_CommercialPaper_pct_sp + (1+sum(ABS_CS_6Mon)/2/100).^0.9*ABS_pct_sp+...
    (1+sum(MBS_CS_6Mon)/2/100).^0.9*AgencyRMBS_pct_sp + (1+sum(CMBS_CS_6Mon)/2/100).^0.9*CMBS_pct_sp -1)};
creditdata(1,2) = {num2str((1+sum(Corp_CS_1Yr)/2/100).^0.9*CorpBond_CommercialPaper_pct_sp + (1+sum(ABS_CS_1Yr)/2/100).^0.9*ABS_pct_sp+...
    (1+sum(MBS_CS_1Yr)/2/100).^0.9*AgencyRMBS_pct_sp + (1+sum(CMBS_CS_1Yr)/2/100).^0.9*CMBS_pct_sp -1)};
creditdata(1,3) = {num2str((1+sum(Corp_CS_2Yr)/2/100).^0.9*CorpBond_CommercialPaper_pct_sp + (1+sum(ABS_CS_2Yr)/2/100).^0.9*ABS_pct_sp+...
    (1+sum(MBS_CS_2Yr)/2/100).^0.9*AgencyRMBS_pct_sp + (1+sum(CMBS_CS_2Yr)/2/100).^0.9*CMBS_pct_sp -1)};
creditdata(1,4) = {num2str((1+sum(Corp_CS_3Yr)/2/100).^0.9*CorpBond_CommercialPaper_pct_sp + (1+sum(ABS_CS_3Yr)/2/100).^0.9*ABS_pct_sp+...
    (1+sum(MBS_CS_3Yr)/2/100).^0.9*AgencyRMBS_pct_sp + (1+sum(CMBS_CS_3Yr)/2/100).^0.9*CMBS_pct_sp -1)};
creditdata(1,5) = {num2str((1+sum(Corp_CS_5Yr)/2/100).^0.9*CorpBond_CommercialPaper_pct_sp + (1+sum(ABS_CS_5Yr)/2/100).^0.9*ABS_pct_sp+...
    (1+sum(MBS_CS_5Yr)/2/100).^0.9*AgencyRMBS_pct_sp + (1+sum(CMBS_CS_5Yr)/2/100).^0.9*CMBS_pct_sp -1)};

n_Corp_CS_6Mon = str2double(get(handles.corp62,'string')) - str2double(get(handles.corp6,'string'));
n_ABS_CS_6Mon = str2double(get(handles.abs62,'string')) - str2double(get(handles.abs6,'string'));
n_MBS_CS_6Mon = str2double(get(handles.agency62,'string')) - str2double(get(handles.agency6,'string'));
n_CMBS_CS_6Mon = str2double(get(handles.cmbs62,'string')) - str2double(get(handles.cmbs6,'string'));

n_Corp_CS_1Yr = str2double(get(handles.corp12,'string')) - str2double(get(handles.corp1,'string'));
n_ABS_CS_1Yr =  str2double(get(handles.abs12,'string')) - str2double(get(handles.abs1,'string'));
n_MBS_CS_1Yr = str2double(get(handles.agency12,'string')) - str2double(get(handles.agency1,'string'));
n_CMBS_CS_1Yr = str2double(get(handles.cmbs12,'string')) - str2double(get(handles.cmbs1,'string'));

n_Corp_CS_2Yr =  str2double(get(handles.corp22,'string')) - str2double(get(handles.corp2,'string'));
n_ABS_CS_2Yr =  str2double(get(handles.abs22,'string')) - str2double(get(handles.abs2,'string'));
n_MBS_CS_2Yr =  str2double(get(handles.agency22,'string')) - str2double(get(handles.agency2,'string'));
n_CMBS_CS_2Yr =  str2double(get(handles.cmbs22,'string')) - str2double(get(handles.cmbs2,'string'));

n_Corp_CS_3Yr =  str2double(get(handles.corp32,'string')) - str2double(get(handles.corp3,'string'));
n_ABS_CS_3Yr =  str2double(get(handles.abs32,'string')) - str2double(get(handles.abs3,'string'));
n_MBS_CS_3Yr =  str2double(get(handles.agency32,'string')) - str2double(get(handles.agency3,'string'));
n_CMBS_CS_3Yr = str2double(get(handles.cmbs32,'string')) -  str2double(get(handles.cmbs3,'string'));

n_Corp_CS_5Yr =  str2double(get(handles.corp52,'string')) - str2double(get(handles.corp5,'string'));
n_ABS_CS_5Yr =  str2double(get(handles.abs52,'string')) - str2double(get(handles.abs5,'string'));
n_MBS_CS_5Yr =  str2double(get(handles.agency52,'string')) - str2double(get(handles.agency5,'string'));
n_CMBS_CS_5Yr =  str2double(get(handles.cmbs52,'string')) - str2double(get(handles.cmbs5,'string'));


sp6 = str2double(get(handles.sp6,'string'));
sp1 = str2double(get(handles.sp1,'string'));
sp2 = str2double(get(handles.sp2,'string'));
sp3 = str2double(get(handles.sp3,'string'));
sp5 = str2double(get(handles.sp5,'string'));

creditdata(2,1) = {num2str(-(n_Corp_CS_6Mon)*sp6/100*CorpBond_CommercialPaper_pct_sp - (n_ABS_CS_6Mon)*sp6/100*ABS_pct_sp...
    -(n_MBS_CS_6Mon)*sp6/100*AgencyRMBS_pct_sp - (n_CMBS_CS_6Mon)*sp6/100*CMBS_pct_sp)};
creditdata(2,2) = {num2str(-(n_Corp_CS_1Yr)*sp1/100*CorpBond_CommercialPaper_pct_sp -(n_ABS_CS_1Yr)*sp1/100*ABS_pct_sp...
    -(n_MBS_CS_1Yr)*sp1/100*AgencyRMBS_pct_sp-(n_CMBS_CS_1Yr)*sp1/100*CMBS_pct_sp)};
creditdata(2,3) = {num2str(-(n_Corp_CS_2Yr)*sp2/100*CorpBond_CommercialPaper_pct_sp -(n_ABS_CS_2Yr)*sp2/100*ABS_pct_sp...
    -(n_MBS_CS_2Yr)*sp2/100*AgencyRMBS_pct_sp-(n_CMBS_CS_2Yr)*sp2/100*CMBS_pct_sp)};
creditdata(2,4) = {num2str(-(n_Corp_CS_3Yr)*sp3/100*CorpBond_CommercialPaper_pct_sp -(n_ABS_CS_3Yr)*sp3/100*ABS_pct_sp...
    -(n_MBS_CS_3Yr)*sp3/100*AgencyRMBS_pct_sp-(n_CMBS_CS_3Yr)*sp3/100*CMBS_pct_sp)};
creditdata(2,5) = {num2str(-(n_Corp_CS_5Yr)*sp5/100*CorpBond_CommercialPaper_pct_sp -(n_ABS_CS_5Yr)*sp5/100*ABS_pct_sp...
    -(n_MBS_CS_5Yr)*sp5/100*AgencyRMBS_pct_sp-(n_CMBS_CS_5Yr)*sp5/100*CMBS_pct_sp)};

creditdata(3,1) = {num2str(str2double(creditdata(2,1)) + str2double(creditdata(1,1)))};
creditdata(3,2) = {num2str(str2double(creditdata(2,2)) + str2double(creditdata(1,2)))};
creditdata(3,3) = {num2str(str2double(creditdata(2,3)) + str2double(creditdata(1,3)))};
creditdata(3,4) = {num2str(str2double(creditdata(2,4)) + str2double(creditdata(1,4)))};
creditdata(3,5) = {num2str(str2double(creditdata(2,5)) + str2double(creditdata(1,5)))};

creditdata(4,1) = {num2str(str2double(intdata(3,1)) + str2double(creditdata(3,1)))};
creditdata(4,2) = {num2str(str2double(intdata(3,2)) + str2double(creditdata(3,2)))};
creditdata(4,3) = {num2str(str2double(intdata(3,3)) + str2double(creditdata(3,3)))};
creditdata(4,4) = {num2str(str2double(intdata(3,4)) + str2double(creditdata(3,4)))};
creditdata(4,5) = {num2str(str2double(intdata(3,5)) + str2double(creditdata(3,5)))};

set(handles.credittable,'data',creditdata);


% --- Executes on button press in final.
function final_Callback(hObject, eventdata, handles)
% hObject    handle to final (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
creditdata = get(handles.credittable,'data');
intdata = get(handles.inttable,'data');
resultdata = get(handles.resulttable,'data');

resultdata (1,1) = {num2str(str2double(intdata(1,1)) + str2double(creditdata(1,1)))};
resultdata (1,2) = {num2str(str2double(intdata(1,2)) + str2double(creditdata(1,2)))};
resultdata (1,3) = {num2str(str2double(intdata(1,3)) + str2double(creditdata(1,3)))};
resultdata (1,4) = {num2str(str2double(intdata(1,4)) + str2double(creditdata(1,4)))};
resultdata (1,5) = {num2str(str2double(intdata(1,5)) + str2double(creditdata(1,5)))};

resultdata (2,1) = {num2str(str2double(intdata(2,1)) + str2double(creditdata(2,1)))};
resultdata (2,2) = {num2str(str2double(intdata(2,2)) + str2double(creditdata(2,2)))};
resultdata (2,3) = {num2str(str2double(intdata(2,3)) + str2double(creditdata(2,3)))};
resultdata (2,4) = {num2str(str2double(intdata(2,4)) + str2double(creditdata(2,4)))};
resultdata (2,5) = {num2str(str2double(intdata(2,5)) + str2double(creditdata(2,5)))};

resultdata (3,1) = {num2str(str2double(resultdata(1,1)) + str2double(resultdata(2,1)))};
resultdata (3,2) = {num2str(str2double(resultdata(1,2)) + str2double(resultdata(2,2)))};
resultdata (3,3) = {num2str(str2double(resultdata(1,3)) + str2double(resultdata(2,3)))};
resultdata (3,4) = {num2str(str2double(resultdata(1,4)) + str2double(resultdata(2,4)))};
resultdata (3,5) = {num2str(str2double(resultdata(1,5)) + str2double(resultdata(2,5)))};

set(handles.resulttable,'data',resultdata);
