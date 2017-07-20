[~, ~, raw] = xlsread('C:\Users\wenbo\Desktop\Amy_Project\regressionData.xlsx','Sheet1','A2:H488');

data = reshape([raw{:}],size(raw));

Mo1_r = data(:,1);
Yr10_r = data(:,2);
Mo3_r = data(:,3);
Mo6_r = data(:,4);
Yr1_r = data(:,5);
Yr2_r = data(:,6);
Yr3_r = data(:,7);
Yr5_r = data(:,8);

clearvars data raw;


InterestRateImpact = cell2table(cell(3,5),'RowNames',{'Income Return','Capital appreciation','Total Rates Impact'},'VariableNames',{'Mon6','Yr1','Yr2','Yr3','Yr5'});
CreditSpreadImpact = cell2table(cell(4,5),'RowNames',{'Income Return','Capital appreciation','Total Credit Impact','Total Return'},'VariableNames',{'Mon6','Yr1','Yr2','Yr3','Yr5'});
result = cell2table(cell(3,5),'RowNames',{'Income Return','Capital appreciation','Total Return'},'VariableNames',{'Mon6','Yr1','Yr2','Yr3','Yr5'});
Agency_RMBS_pct = 0.07;
Agency_RMBS_DurationExtension = 1.3;

CorpBond_CommercialPaper_pct_sp = 0.64;
ABS_pct_sp = 0.22;
AgencyRMBS_pct_sp = 0.08;
CMBS_pct_sp = 0.06;

Date = {'2/6/2017'; '3/31/2017';'6/30/2017';'9/30/2017';'12/31/2017'};
fed_fund_rate = [0.625;0.875;0.875;1.125;1.125];
Mo1 = [0.41;0.88;0.88;1.13;1.13];

Mo3 = zeros(5,1);
Mo3(1) = [0.54];

Mo6 = zeros(5,1);
Mo6(1) = [0.64];

Yr1 = zeros(5,1);
Yr1(1) = [0.85];

Yr2 = zeros(5,1);
Yr2(1) = [1.15];

Yr3 = zeros(5,1);
Yr3(1) = [1.43];

Yr5  = zeros(5,1);
Yr5(1) = [1.89];

Yr10 = zeros(5,1);
Yr10(1) = 2.47;
Yr10(5) = 3;

para = regress(Mo3_r,[Mo1_r,Yr10_r,ones(size(Mo1_r))]);
Mo3(5) = para(1)*1.13 + para(2)*3+para(3);

para = regress(Mo6_r,[Mo1_r,Yr10_r,ones(size(Mo1_r))]);
Mo6(5) = para(1)*1.13 + para(2)*3+para(3);

para = regress(Yr1_r,[Mo1_r,Yr10_r,ones(size(Mo1_r))]);
Yr1(5) = para(1)*1.13 + para(2)*3+para(3);

para = regress(Yr2_r,[Mo1_r,Yr10_r,ones(size(Mo1_r))]);
Yr2(5) = para(1)*1.13 + para(2)*3+para(3);

para = regress(Yr3_r,[Mo1_r,Yr10_r,ones(size(Mo1_r))]);
Yr3(5) = para(1)*1.13 + para(2)*3+para(3);

para = regress(Yr5_r,[Mo1_r,Yr10_r,ones(size(Mo1_r))]);
Yr5(5) = para(1)*1.13 + para(2)*3+para(3);

for i = 2:5
    Mo3(i) = (Mo3(5)-Mo3(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Mo3(1);
    Mo6(i) = (Mo6(5)-Mo6(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Mo6(1);
    Yr1(i) = (Yr1(5)-Yr1(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Yr1(1);
    Yr2(i) = (Yr2(5)-Yr2(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Yr2(1);
    Yr3(i) = (Yr3(5)-Yr3(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Yr3(1);
    Yr5(i) = (Yr5(5)-Yr5(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Yr5(1);
    Yr10(i) = (Yr10(5)-Yr10(1))*(datenum(Date(i))-datenum(Date(1)))/(datenum(Date(5))-datenum(Date(1)))+Yr10(1);
end

T1 = table(fed_fund_rate,Mo1,Mo3,Mo6,Yr1,Yr2,Yr3,Yr5,Yr10,'RowNames',Date);

Date2 = {'2/6/2017';'3/31/2017';'4/30/2017';'5/31/2017';'6/30/2017';'7/31/2017';'8/31/2017';'9/30/2017';'10/31/2017';'11/30/2017';'12/31/2017'};
Mo1_2 = [0.41;0.88;0.875;0.875;0.88;0.958333;1.0416667;1.13;1.125;1.125;1.13];
Mo3_2 = [0.54;0.69;0.77516;0.86076;0.95;1.03289;1.11943;1.21;1.29251;1.37905;1.47];
Mo6_2 = [0.64;0.85;0.965828;1.084429;1.2;1.322935;1.442839;1.56;1.682648;1.802553;1.92];
Yr1_2 = [0.85;1.06;1.180444;1.300725;1.42;1.54261;1.664214;1.79;1.90742;2.029024;2.15];
Yr2_2 = [1.15;1.34;1.442682;1.549218;1.66;1.763462;1.871168;1.98;2.086582;;2.194289;2.3];
Yr3_2 = [1.43;1.58;1.659003;1.74236;1.83;1.90999;1.994263;2.08;2.162809;2.247082;2.33];
Yr5_2 = [1.89;1.99;2.043945;2.09998;2.16;2.212668;2.269319;2.33;2.382623;2.439274;2.5];

T2 = table(Mo1_2,Mo3_2,Mo6_2,Yr1_2,Yr2_2,Yr3_2,Yr5_2,'RowNames',Date2);

income_col_6mo_1 = 1 + Mo6_2(1)/400*1;
income_col_6mo_2 = income_col_6mo_1*0.75*(1+ Mo6_2(1)/400)+ income_col_6mo_1*0.25*(1+Mo6_2(4)/400);
income_col_6mo_3 = income_col_6mo_2*0.5*(1 + Mo6_2(1)/400)+ income_col_6mo_2*0.25*(1+Mo6_2(4)/400) + income_col_6mo_2*0.25*(1+Mo6_2(7)/400);
income_col_6mo_final = income_col_6mo_3*0.25*(1 + Mo6_2(1)/400)+ income_col_6mo_3*0.25*(1+Mo6_2(4)/400) + income_col_6mo_3*0.25*(1+Mo6_2(7)/400)+ income_col_6mo_3*0.25*(1+Mo6_2(10)/400);
InterestRateImpact.Mon6{'Income Return'} = income_col_6mo_final - 1;

income_col_1yr_1 = 1 + Yr1_2(1)/400*1;
income_col_1yr_2 = income_col_1yr_1*0.8*(1+ Yr1_2(1)/400)+ income_col_1yr_1*0.2*(1+Yr1_2(4)/400);
income_col_1yr_3 = income_col_1yr_2*0.6*(1 + Yr1_2(1)/400)+ income_col_1yr_2*0.2*(1+Yr1_2(4)/400) + income_col_1yr_2*0.2*(1+Yr1_2(7)/400);
income_col_1yr_final = income_col_1yr_3*0.4*(1 + Yr1_2(1)/400)+ income_col_1yr_3*0.2*(1+Yr1_2(4)/400) + income_col_1yr_3*0.2*(1+Yr1_2(7)/400)+ income_col_1yr_3*0.2*(1+Yr1_2(10)/400);
InterestRateImpact.Yr1{'Income Return'} = income_col_1yr_final - 1;

income_col_2yr_1 = 1 + Yr2_2(1)/400*1;
income_col_2yr_2 = income_col_2yr_1*0.875*(1+ Yr2_2(1)/400)+ income_col_2yr_1*0.125*(1+Yr2_2(4)/400);
income_col_2yr_3 = income_col_2yr_2*0.75*(1 + Yr2_2(1)/400)+ income_col_2yr_2*0.125*(1+Yr2_2(4)/400) + income_col_2yr_2*0.125*(1+Yr2_2(7)/400);
income_col_2yr_final = income_col_2yr_3*0.625*(1 + Yr2_2(1)/400)+ income_col_2yr_3*0.125*(1+Yr2_2(4)/400) + income_col_2yr_3*0.125*(1+Yr2_2(7)/400)+ income_col_2yr_3*0.125*(1+Yr2_2(10)/400);
InterestRateImpact.Yr2{'Income Return'} = income_col_2yr_final - 1;

income_col_3yr_1 = 1 + Yr3_2(1)/400*1;
income_col_3yr_2 = income_col_3yr_1*0.92*(1+ Yr3_2(1)/400)+ income_col_3yr_1*0.08*(1+Yr3_2(4)/400);
income_col_3yr_3 = income_col_3yr_2*0.84*(1 + Yr3_2(1)/400)+ income_col_3yr_2*0.08*(1+Yr3_2(4)/400) + income_col_3yr_2*0.08*(1+Yr3_2(7)/400);
income_col_3yr_final = income_col_3yr_3*0.76*(1 + Yr3_2(1)/400)+ income_col_3yr_3*0.08*(1+Yr3_2(4)/400) + income_col_3yr_3*0.08*(1+Yr3_2(7)/400)+ income_col_3yr_3*0.08*(1+Yr3_2(10)/400);
InterestRateImpact.Yr3{'Income Return'} = income_col_3yr_final - 1;

income_col_5yr_1 = 1 + Yr5_2(1)/400*1;
income_col_5yr_2 = income_col_5yr_1*0.95*(1+ Yr5_2(1)/400)+ income_col_5yr_1*0.05*(1+Yr5_2(4)/400);
income_col_5yr_3 = income_col_5yr_2*0.9*(1 + Yr5_2(1)/400)+ income_col_5yr_2*0.05*(1+Yr5_2(4)/400) + income_col_5yr_2*0.05*(1+Yr5_2(7)/400);
income_col_5yr_final = income_col_5yr_3*0.85*(1 + Yr5_2(1)/400)+ income_col_5yr_3*0.05*(1+Yr5_2(4)/400) + income_col_5yr_3*0.05*(1+Yr5_2(7)/400)+ income_col_5yr_3*0.05*(1+Yr5_2(10)/400);
InterestRateImpact.Yr5{'Income Return'} = income_col_5yr_final - 1;



Mo3_d = [0.246;0.85];
Mo6_d = [0.494;1.0075];
Yr1_d = [0.99;1.3225];
Yr2_d = [1.919;1.9525];
Yr3_d = [2.928;2.778333];
Yr5_d = [4.719;4.43];
duration = {'rate Modeified Duration','Spread Duration'};
T3 = table(Mo3_d,Mo6_d,Yr1_d,Yr2_d,Yr3_d,Yr5_d,'RowNames',duration);

InterestRateImpact.Mon6{'Capital appreciation'}=(Mo6(5)-Mo6(1))*Mo6_d(1)/-100*(1-Agency_RMBS_pct)+(Mo6(5)-Mo6(1))*Mo6_d(1)/-100*Agency_RMBS_pct*Agency_RMBS_DurationExtension;
InterestRateImpact.Yr1{'Capital appreciation'}=(Yr1(5)-Yr1(1))*Yr1_d(1)/-100*(1-Agency_RMBS_pct)+(Yr1(5)-Yr1(1))*Yr1_d(1)/-100*Agency_RMBS_pct*Agency_RMBS_DurationExtension;
InterestRateImpact.Yr2{'Capital appreciation'}=(Yr2(5)-Yr2(1))*Yr2_d(1)/-100*(1-Agency_RMBS_pct)+(Yr2(5)-Yr2(1))*Yr2_d(1)/-100*Agency_RMBS_pct*Agency_RMBS_DurationExtension;
InterestRateImpact.Yr3{'Capital appreciation'}=(Yr3(5)-Yr3(1))*Yr3_d(1)/-100*(1-Agency_RMBS_pct)+(Yr3(5)-Yr3(1))*Yr3_d(1)/-100*Agency_RMBS_pct*Agency_RMBS_DurationExtension;
InterestRateImpact.Yr5{'Capital appreciation'}=(Yr5(5)-Yr5(1))*Yr5_d(1)/-100*(1-Agency_RMBS_pct)+(Yr5(5)-Yr5(1))*Yr5_d(1)/-100*Agency_RMBS_pct*Agency_RMBS_DurationExtension;


InterestRateImpact.Mon6{'Total Rates Impact'} = InterestRateImpact.Mon6{'Capital appreciation'}+InterestRateImpact.Mon6{'Income Return'};
InterestRateImpact.Yr1{'Total Rates Impact'} = InterestRateImpact.Yr1{'Capital appreciation'}+InterestRateImpact.Yr1{'Income Return'};
InterestRateImpact.Yr2{'Total Rates Impact'} = InterestRateImpact.Yr2{'Capital appreciation'}+InterestRateImpact.Yr2{'Income Return'};
InterestRateImpact.Yr3{'Total Rates Impact'} = InterestRateImpact.Yr3{'Capital appreciation'}+InterestRateImpact.Yr3{'Income Return'};
InterestRateImpact.Yr5{'Total Rates Impact'} = InterestRateImpact.Yr5{'Capital appreciation'}+InterestRateImpact.Yr5{'Income Return'}

Corp_CS_6Mon = [0.38;0.5];
Corp_CS_1Yr = [0.76;1.01];
Corp_CS_2Yr = [0.85;0.96];
Corp_CS_3Yr = [0.94;0.92];
Corp_CS_5Yr = [1.12;0.94];

ABS_CS_6Mon = [0.2;0.21];
ABS_CS_1Yr = [0.4;0.41];
ABS_CS_2Yr = [0.47;0.39];
ABS_CS_3Yr = [0.53;0.37];
ABS_CS_5Yr = [1.08;0.72];

MBS_CS_6Mon = [0.7;-0.25];
MBS_CS_1Yr = [1.4;-0.49];
MBS_CS_2Yr = [1.03;-0.36];
MBS_CS_3Yr = [0.66;-0.23];
MBS_CS_5Yr = [0.26;-0.04];

CMBS_CS_6Mon = [0.32;-0.01];
CMBS_CS_1Yr = [0.65;-0.02];
CMBS_CS_2Yr = [0.58;0.07];
CMBS_CS_3Yr = [0.51;0.16];
CMBS_CS_5Yr = [0.5;0.35];


CreditSpreadImpact.Mon6{'Income Return'} = (1+sum(Corp_CS_6Mon)/2/100).^0.9*CorpBond_CommercialPaper_pct_sp + (1+sum(ABS_CS_6Mon)/2/100).^0.9*ABS_pct_sp+...
    (1+sum(MBS_CS_6Mon)/2/100).^0.9*AgencyRMBS_pct_sp + (1+sum(CMBS_CS_6Mon)/2/100).^0.9*CMBS_pct_sp -1;
CreditSpreadImpact.Yr1{'Income Return'} = (1+sum(Corp_CS_1Yr)/2/100).^0.9*CorpBond_CommercialPaper_pct_sp + (1+sum(ABS_CS_1Yr)/2/100).^0.9*ABS_pct_sp+...
    (1+sum(MBS_CS_1Yr)/2/100).^0.9*AgencyRMBS_pct_sp + (1+sum(CMBS_CS_1Yr)/2/100).^0.9*CMBS_pct_sp -1;
CreditSpreadImpact.Yr2{'Income Return'} = (1+sum(Corp_CS_2Yr)/2/100).^0.9*CorpBond_CommercialPaper_pct_sp + (1+sum(ABS_CS_2Yr)/2/100).^0.9*ABS_pct_sp+...
    (1+sum(MBS_CS_2Yr)/2/100).^0.9*AgencyRMBS_pct_sp + (1+sum(CMBS_CS_2Yr)/2/100).^0.9*CMBS_pct_sp -1;
CreditSpreadImpact.Yr3{'Income Return'} = (1+sum(Corp_CS_3Yr)/2/100).^0.9*CorpBond_CommercialPaper_pct_sp + (1+sum(ABS_CS_3Yr)/2/100).^0.9*ABS_pct_sp+...
    (1+sum(MBS_CS_3Yr)/2/100).^0.9*AgencyRMBS_pct_sp + (1+sum(CMBS_CS_3Yr)/2/100).^0.9*CMBS_pct_sp -1;
CreditSpreadImpact.Yr5{'Income Return'} = (1+sum(Corp_CS_5Yr)/2/100).^0.9*CorpBond_CommercialPaper_pct_sp + (1+sum(ABS_CS_5Yr)/2/100).^0.9*ABS_pct_sp+...
    (1+sum(MBS_CS_5Yr)/2/100).^0.9*AgencyRMBS_pct_sp + (1+sum(CMBS_CS_5Yr)/2/100).^0.9*CMBS_pct_sp -1;



CreditSpreadImpact.Mon6{'Capital appreciation'} = -diff(Corp_CS_6Mon)*Mo6_d(2)/100*CorpBond_CommercialPaper_pct_sp - diff(ABS_CS_6Mon)*Mo6_d(2)/100*ABS_pct_sp...
    -diff(MBS_CS_6Mon)*Mo6_d(2)/100*AgencyRMBS_pct_sp-diff(CMBS_CS_6Mon)*Mo6_d(2)/100*CMBS_pct_sp;
CreditSpreadImpact.Yr1{'Capital appreciation'} = -diff(Corp_CS_1Yr)*Yr1_d(2)/100*CorpBond_CommercialPaper_pct_sp - diff(ABS_CS_1Yr)*Yr1_d(2)/100*ABS_pct_sp...
    -diff(MBS_CS_1Yr)*Yr1_d(2)/100*AgencyRMBS_pct_sp-diff(CMBS_CS_1Yr)*Yr1_d(2)/100*CMBS_pct_sp;
CreditSpreadImpact.Yr2{'Capital appreciation'} = -diff(Corp_CS_2Yr)*Yr2_d(2)/100*CorpBond_CommercialPaper_pct_sp - diff(ABS_CS_2Yr)*Yr2_d(2)/100*ABS_pct_sp...
    -diff(MBS_CS_2Yr)*Yr2_d(2)/100*AgencyRMBS_pct_sp-diff(CMBS_CS_2Yr)*Yr2_d(2)/100*CMBS_pct_sp;
CreditSpreadImpact.Yr3{'Capital appreciation'} = -diff(Corp_CS_3Yr)*Yr3_d(2)/100*CorpBond_CommercialPaper_pct_sp - diff(ABS_CS_3Yr)*Yr3_d(2)/100*ABS_pct_sp...
    -diff(MBS_CS_3Yr)*Yr3_d(2)/100*AgencyRMBS_pct_sp-diff(CMBS_CS_3Yr)*Yr3_d(2)/100*CMBS_pct_sp;
CreditSpreadImpact.Yr5{'Capital appreciation'} = -diff(Corp_CS_5Yr)*Yr5_d(2)/100*CorpBond_CommercialPaper_pct_sp - diff(ABS_CS_5Yr)*Yr5_d(2)/100*ABS_pct_sp...
    -diff(MBS_CS_5Yr)*Yr5_d(2)/100*AgencyRMBS_pct_sp-diff(CMBS_CS_5Yr)*Yr5_d(2)/100*CMBS_pct_sp;


CreditSpreadImpact.Mon6{'Total Credit Impact'} = CreditSpreadImpact.Mon6{'Income Return'}+CreditSpreadImpact.Mon6{'Capital appreciation'};
CreditSpreadImpact.Yr1{'Total Credit Impact'} = CreditSpreadImpact.Yr1{'Income Return'}+CreditSpreadImpact.Yr1{'Capital appreciation'};
CreditSpreadImpact.Yr2{'Total Credit Impact'} = CreditSpreadImpact.Yr2{'Income Return'}+CreditSpreadImpact.Yr2{'Capital appreciation'};
CreditSpreadImpact.Yr3{'Total Credit Impact'} = CreditSpreadImpact.Yr3{'Income Return'}+CreditSpreadImpact.Yr3{'Capital appreciation'};
CreditSpreadImpact.Yr5{'Total Credit Impact'} = CreditSpreadImpact.Yr5{'Income Return'}+CreditSpreadImpact.Yr5{'Capital appreciation'};

CreditSpreadImpact.Mon6{'Total Return'} = InterestRateImpact.Mon6{'Capital appreciation'}+InterestRateImpact.Mon6{'Income Return'}...
    + CreditSpreadImpact.Mon6{'Income Return'}+CreditSpreadImpact.Mon6{'Capital appreciation'};
CreditSpreadImpact.Yr1{'Total Return'} = InterestRateImpact.Yr1{'Capital appreciation'}+InterestRateImpact.Yr1{'Income Return'}...
    + CreditSpreadImpact.Yr1{'Income Return'}+CreditSpreadImpact.Yr1{'Capital appreciation'};
CreditSpreadImpact.Yr2{'Total Return'} = InterestRateImpact.Yr2{'Capital appreciation'}+InterestRateImpact.Yr2{'Income Return'}...
    + CreditSpreadImpact.Yr2{'Income Return'}+CreditSpreadImpact.Yr2{'Capital appreciation'};
CreditSpreadImpact.Yr3{'Total Return'} = InterestRateImpact.Yr3{'Capital appreciation'}+InterestRateImpact.Yr3{'Income Return'}...
    + CreditSpreadImpact.Yr3{'Income Return'}+CreditSpreadImpact.Yr3{'Capital appreciation'};
CreditSpreadImpact.Yr5{'Total Return'} = InterestRateImpact.Yr5{'Capital appreciation'}+InterestRateImpact.Yr5{'Income Return'}...
    + CreditSpreadImpact.Yr5{'Income Return'}+CreditSpreadImpact.Yr5{'Capital appreciation'}

result.Mon6{'Income Return'} = InterestRateImpact.Mon6{'Income Return'}+CreditSpreadImpact.Mon6{'Income Return'};
result.Mon6{'Capital appreciation'} = InterestRateImpact.Mon6{'Capital appreciation'}+CreditSpreadImpact.Mon6{'Capital appreciation'};
result.Mon6{'Total Return'} = result.Mon6{'Income Return'}+result.Mon6{'Capital appreciation'};

result.Yr1{'Income Return'} = InterestRateImpact.Yr1{'Income Return'}+CreditSpreadImpact.Yr1{'Income Return'};
result.Yr1{'Capital appreciation'} = InterestRateImpact.Yr1{'Capital appreciation'}+CreditSpreadImpact.Yr1{'Capital appreciation'};
result.Yr1{'Total Return'} = result.Yr1{'Income Return'}+result.Yr1{'Capital appreciation'};

result.Yr2{'Income Return'} = InterestRateImpact.Yr2{'Income Return'}+CreditSpreadImpact.Yr2{'Income Return'};
result.Yr2{'Capital appreciation'} = InterestRateImpact.Yr2{'Capital appreciation'}+CreditSpreadImpact.Yr2{'Capital appreciation'};
result.Yr2{'Total Return'} = result.Yr2{'Income Return'}+result.Yr2{'Capital appreciation'};

result.Yr3{'Income Return'} = InterestRateImpact.Yr3{'Income Return'}+CreditSpreadImpact.Yr3{'Income Return'};
result.Yr3{'Capital appreciation'} = InterestRateImpact.Yr3{'Capital appreciation'}+CreditSpreadImpact.Yr3{'Capital appreciation'};
result.Yr3{'Total Return'} = result.Yr3{'Income Return'}+result.Yr3{'Capital appreciation'};

result.Yr5{'Income Return'} = InterestRateImpact.Yr5{'Income Return'}+CreditSpreadImpact.Yr5{'Income Return'};
result.Yr5{'Capital appreciation'} = InterestRateImpact.Yr5{'Capital appreciation'}+CreditSpreadImpact.Yr5{'Capital appreciation'};
result.Yr5{'Total Return'} = result.Yr5{'Income Return'}+result.Yr5{'Capital appreciation'}

 
