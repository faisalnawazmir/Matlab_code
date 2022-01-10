% mir@ciit-attock.edu.pk, developed on February 05, 2018.
%% Assume no of companies m=26, no of observations n=571 for each country, and no of year t=12. 
%Instructions: All the columns of excel sheet should be in number format.
% 1. contains Date in Date number format,
% 2. contins compnay returns,
% 3. last 5 columns sholud be MI MI_Lag1 MI_Lag2 MI_Lead1 MI_Lead2
% i.e., Market Index, their 2 Leads, and 2 Lags respectively.

clear all
clc
%% 
%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% SETTING THE PATHS
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
data_path = 'C:\Users\Faisal\Documents\MATLAB\crashRisk\';         % where the data spreadsheet is saved
save_path = 'C:\Users\Faisal\Documents\MATLAB\crashRisk\';         % where you would like any output MAT files saved
save_name = 'output_crashriskk';                       % files will be saved as 'save_name_stage_x.mat' for x=1,2,...,14

%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% LOADING IN THE DATA
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% Provide filename sheet name and range of data for import of data purpose.
data = xlsread([data_path,'3.xlsx'],'Sheet1','a2:ES570');

Dateexcel = data(:,1); % date in number
Date = x2mdate(Dateexcel)%Excel serial date number to MATLAB serial date number or datetime format
CD= data(:,2:end-5);
D=data(:,end-4:end);% Market Index, their 2 Leads, and 2 Lags respectively.
%D=[MI MI_Lag1 MI_Lag2 MI_Lead1 MI_Lead2];
%% 
%CD=crashrisk1; % import data from excel and change name to crashrisk 1or 2.
noObs=length(CD);
a=size(CD);
noOfCmp=a(1,2);
noOfyear=12;
d=0;
r=repmat(d,noObs,noOfCmp);% 26 is number of countries, 571 observations for each country.
b=repmat(d,5,1);% no of parameters do not change it.
bint=repmat(d,(noOfCmp-1),2);% no of parameters, do not change it.

%% Extracting Residual from the market model. 
for i=1:noOfCmp
[bi,binti,r(:,i)] = regress(CD(:,i),D)
end

%% transform the residual return
rr=r+1
ret=log(rr);
rret=real(ret);%rret=Wit in paper pp.12 para first.
% http://www.fmaconferences.org/Tokyo/Papers/IT_Crash_WRZhang.pdf
%% 
ret2=rret.^2;
ret3=rret.^3;
%% 
yearsum2=repmat(d,noOfyear,noOfCmp);
yearsum3=repmat(d,noOfyear,noOfCmp);
yearsum=repmat(d,noOfyear,noOfCmp);
yearmean=repmat(d,noOfyear,noOfCmp);
dv = datevec(Date);                   %convert to datevec to easily separate years
[years, ~, subs] = unique(dv(:, 1));     %get unique years and location
%% 
i=0;
for i=1:noOfCmp
yearsum2(:,i) = accumarray(subs, ret2(:, i));     %accumarray with most default values does sums
yearsum3(:,i) = accumarray(subs, ret3(:, i));
yearmean(:,i) = accumarray(subs, rret(:, i),[],@mean);
end


%% 
yearcount = [years, accumarray(subs, 1)];% n week per year
n=yearcount(:,2)% only count from year,count

%% 
%NCSKEW=-[((n.*(n-1)).^(1.5))*yearsum3]/[(n-1).*(n-2).*(yearsum2).^(1.5)]
%sum(~isnan(A),2), sum(A==A,2);
% I have a matrix for which I want to compute the number of non-nan observations 
% for each row without running a loop. is this possible?
%% 
A=repmat(d,noOfyear,noOfCmp);
j=0;

for j=1:noOfCmp
A(:,j)=[n.*((n-1).^1.5).*(yearsum3(:,j))]./[((n-1).*(n-2)).*((yearsum2(:,j)).^1.5)]
end
NCSKEW=-A;
%% work for DUVOL
i=0;
for i=1:noOfCmp
avgm(:,i)=yearmean(subs,i);
end
%% 
m=repmat(d,noObs,noOfCmp)
i=0;
for i=1:noOfCmp
m(:,i)=rret(:,i)<avgm(:,i)
end
%%  
Dret=repmat(d,noObs,noOfCmp);
Uret=repmat(d,noObs,noOfCmp); 
subsdownrett=repmat(d,noObs,noOfCmp);
subsuprett=repmat(d,noObs,noOfCmp);

%% 
for i=1:noOfCmp
  subsdownrett(:,i)=(m(:,i)==1);
  subsuprett(:,i)=(m(:,i)==0);
  end

%% 
i=0;
for i=1:noOfCmp
Dret(:,i)=subsdownrett(:,i).* rret(:,i);
Uret(:,i)=subsuprett(:,i).* rret(:,i);
end  

%% 
Dret2=Dret.^2;   
Uret2=Uret.^2;
Uret2sum=repmat(d,noOfyear,noOfCmp);
Dret2sum=repmat(d,noOfyear,noOfCmp);
n_d=repmat(d,noOfyear,noOfCmp);
n_u=repmat(d,noOfyear,noOfCmp);

%% 
i=0;
for i=1:noOfCmp
Uret2sum(:,i) = accumarray(subs, Uret2(:, i));
Dret2sum(:,i) = accumarray(subs, Dret2(:, i));
n_d(:,i)=accumarray(subs, subsdownrett(:,i));
n_u(:,i)=accumarray(subs, subsuprett(:,i));
end
%% 
z_d=n_d-1;
z_u=n_u-1;

%%  
i=0;
for i=1:noOfCmp
Ustd(:,i)=[Uret2sum(:,i)].*[z_d(:,i)]
Dstd(:,i)=[Dret2sum(:,i)].*[z_u(:,i)]
end

%%
DUVOL=log([Dstd] ./[Ustd]);
%% write year
filename = 'crashriskresultcountry3.xlsx'; % change 1,2 3, like 
A = [2006:2017]';
sheet = 1;
xlRange = 'A2';
xlswrite(filename,A,sheet,xlRange)

%% write 1 to companies no say 29
AA = [1:noOfCmp];
B=repmat(AA,1,2);
sheet = 1;
xlRange = 'B1';
xlswrite(filename,B,sheet,xlRange)
%% write two measure for each year each company
C=[NCSKEW DUVOL];
sheet = 1;
xlRange = 'B2';
xlswrite(filename,C,sheet,xlRange)
%% 



