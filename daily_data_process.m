%% RISK REPORT GENERATION AUTOMATION
%  AUTHOR: YORK YU
%  SUPERVISOR: ROB JACKSON
%  INTERNAL USE ONLY
%  PURPOSE: TO AUTOMATE RISK REPORT FOR DIFFERENT FUNDS

%% SET TIME AND LOAD IN NAV REPORT AND DATA
% load time
Time = '02/23/2015';
% load file directory
file = 'C:\Users\yyu\Desktop\Cyrus Consolidated Risk Report\CSSM_022315_Report.xlsx';
% load NAV record
NAV = xlsread('C:\Users\yyu\Desktop\Cyrus Consolidated Risk Report\NAV\CSSM_NAV_2015.02.23.xlsx');
% load main file
[num,text,first]= xlsread('C:\Users\yyu\Desktop\Cyrus Consolidated Risk Report\Position\CSSM_022315.xls');
raw = first(:,[10 17 18 22 26 6 23 12 13 15 31 32]);
%% YTD DATA - FROM NAV REPORT
YTD = zeros(18,2);
navn = NAV(end); % current NAV
NAV_length = length(NAV); % number of NAV records
YTD(1,1) = navn - NAV(end-1); % Total last day PNL
YTD(1,2) = navn/NAV(end-1) - 1;
NAV_YEAR_START = 343638657;
NAV_MONTH_START = 345534595;
YTD(2,1) = navn - NAV_YEAR_START; % YTD PNL
YTD(2,2) = navn/NAV_YEAR_START - 1; % 
if NAV_length<21
YTD(3,1) = navn - NAV(1); % MTD PNL
YTD(3,2) = navn/NAV(1) - 1;
else
YTD(3,1) = navn - NAV(end-20); % MTD PNL
YTD(3,2) = navn/NAV(end-20) - 1;
end
YTD(4,1) = YTD(2,1)*(252/NAV_length); % Annualized PNL from YTD PNL
YTD(4,2) = (YTD(2,2)+1)^(252/NAV_length)-1;
YTD(5,1) = std(NAV); %Volatility Year to Date
% find return
NAV_return = zeros(NAV_length,1);
for i = 2: NAV_length
    NAV_return(i) = NAV(i)/NAV(i-1)-1;
end
YTD(5,2) = std(NAV_return);
YTD(6,1) = YTD(5,1)*sqrt(252/NAV_length); % Annualized Volatility
YTD(6,2) = YTD(5,2)*sqrt(252/NAV_length); 
YTD(7,1) = NAV_length; % No of days
YTD(7,2) = 1;
YTD(8,1) = NAV_length/2 + (nansum(NAV_return./abs(NAV_return))-sum(isinf(1./NAV_return)))/2; % No of positive days
YTD(8,2) = YTD(8,1)/YTD(7,1);
YTD(9,1) = NAV_length-YTD(8,1); % No of negative days
YTD(9,2) = YTD(9,1)/YTD(7,1);
% Sharpe Ratio
YTD(10,1) = YTD(4,1)/YTD(6,1);
% downside deviation
NAV_difference = NAV(2:end) - NAV(1:end-1);
YTD(11,2) = std((NAV_return-abs(NAV_return))/2);
YTD(11,1) = YTD(3,1)/YTD(3,2)*YTD(11,2);
YTD(12,1) = YTD(11,1)*(252/NAV_length); 
YTD(12,2) = YTD(11,2)*sqrt(252/NAV_length); 
% Sortino Ratio
SR = 0;
count = 0;
target_return = 5.5476e-04;
for i = 1:length(NAV_return)
    if NAV_return(i) < target_return
        count = count + 1;
        SR = SR + (NAV_return(i) - target_return)^2;
    end
end
YTD(13,1) = mean(NAV_return)/(SR/count);
% Max gain and draw
YTD(14,1) = max(NAV_difference);
YTD(14,2) = YTD(14,1)/navn;
YTD(15,1) = max(NAV_difference) - min(NAV_difference);
YTD(15,2) = YTD(15,1)/navn;
YTD(16,1) = mean(NAV_difference)/YTD(15,1);
YTD(17,1) = NAV(end); % AUM
YTD(18,2) = 0.1071699; % YTD SPX Index Change % NEED TO BE ADJUSTED

%% DTD DATA
%  merge positions
proc = raw(2:end,:);
i = 1;
while i ~= size(proc,1)
    i = i+1;
    if length(proc{i,1}) ~= length(proc{i-1,1})
        continue
    end
        if all(proc{i,1} == proc{i-1,1})
            proc{i-1,2} = proc{i-1,2} + proc{i,2};
            proc{i-1,10} = proc{i-1,10} + proc{i,10};
            proc(i,:) = [];
            i = i-1;
        end
end

%%  Top long/short exposures
temp = sortcell(proc,[2 2]);
%  top twenty long
TTL = cell(20,7);
%  top twenty short
TTS = cell(20,7); 

TTL(:,1) = temp(end-19:end,1);
TTL(:,2) = temp(end-19:end,2);
TTL(:,3) = num2cell(cell2mat(TTL(:,2))/navn);
TTL(:,4) = temp(end-19:end,8);
TTL(:,5) = temp(end-19:end,7);
TTL(:,6) = temp(end-19:end,10);
TTL(:,7) = TTL(:,2);
TTL = flipud(TTL);

TTS(:,1) = temp(1:20,1);
TTS(:,2) = temp(1:20,2);
TTS(:,3) = num2cell(cell2mat(TTS(:,2))/navn);
TTS(:,4) = temp(1:20,8);
TTS(:,5) = temp(1:20,7);
TTS(:,6) = temp(1:20,10);
TTS(:,7) = TTS(:,2);

%%  Exposure by fund 
EBF = cell(1,4);
temp = cell2mat(raw(2:end,2));
EBF{1,1} = (sum(temp)+sum(abs(temp)))/2;
EBF{1,2} = (sum(temp)-sum(abs(temp)))/2;
EBF{1,3} = EBF{1,1} - EBF{1,2};
EBF{1,4} = EBF{1,1} + EBF{1,2};
%  percentage
EBFR = num2cell(cell2mat(EBF)/navn);

%%   Exposure by industry
temp = sortcell(proc,[3 2]);
EBI = num2cell(zeros(14,4));
for i = 1:size(temp,1)
    switch temp{i,3}
        case 'Basic Materials'
            EBI{1,3} = EBI{1,3} + abs(temp{i,2});
            EBI{1,4} = EBI{1,4} + temp{i,2};
        case 'Communications'
            EBI{2,3} = EBI{2,3} + abs(temp{i,2});
            EBI{2,4} = EBI{2,4} + temp{i,2};            
        case 'Consumer, Cyclical'
            EBI{3,3} = EBI{3,3} + abs(temp{i,2});
            EBI{3,4} = EBI{3,4} + temp{i,2};          
        case 'Consumer, Non-cyclical'
            EBI{4,3} = EBI{4,3} + abs(temp{i,2});
            EBI{4,4} = EBI{4,4} + temp{i,2};            
        case 'Diversified'
            EBI{5,3} = EBI{5,3} + abs(temp{i,2});
            EBI{5,4} = EBI{5,4} + temp{i,2};            
        case 'Energy'
            EBI{6,3} = EBI{6,3} + abs(temp{i,2});
            EBI{6,4} = EBI{6,4} + temp{i,2};            
        case 'Financial'
            EBI{7,3} = EBI{7,3} + abs(temp{i,2});
            EBI{7,4} = EBI{7,4} + temp{i,2};            
        case 'Health Care'
            EBI{8,3} = EBI{8,3} + abs(temp{i,2});
            EBI{8,4} = EBI{8,4} + temp{i,2};
        case 'Hedges/Non-classified'
            EBI{9,3} = EBI{9,3} + abs(temp{i,2});
            EBI{9,4} = EBI{9,4} + temp{i,2};          
        case 'Industrial'
            EBI{10,3} = EBI{10,3} + abs(temp{i,2});
            EBI{10,4} = EBI{10,4} + temp{i,2};            
        case 'Sovereign / Government'
            EBI{11,3} = EBI{11,3} + abs(temp{i,2});
            EBI{11,4} = EBI{11,4} + temp{i,2};            
        case 'Technology'
            EBI{12,3} = EBI{12,3} + abs(temp{i,2});
            EBI{12,4} = EBI{12,4} + temp{i,2};
        case 'Utilities'
            EBI{13,3} = EBI{13,3} + abs(temp{i,2});
            EBI{13,4} = EBI{13,4} + temp{i,2};
    end
end
for i = 1:13
    EBI{i,1} = (EBI{i,3}+EBI{i,4})/2;
    EBI{i,2} = (-EBI{i,3}+EBI{i,4})/2;
end
EBI{14,1} = sum(cell2mat(EBI(1:13,1)));
EBI{14,2} = sum(cell2mat(EBI(1:13,2)));
EBI{14,3} = sum(cell2mat(EBI(1:13,3)));
EBI{14,4} = sum(cell2mat(EBI(1:13,4)));
%  percentage
EBIR = num2cell(cell2mat(EBI)/navn);

%% Exposure by strat
temp = sortcell(proc,[4 2]);
EBS = num2cell(zeros(7,4));
for i = 1:size(temp,1)
    switch temp{i,4}
        case 'Currency Hedges'
            EBS{1,3} = EBS{1,3} + abs(temp{i,2});
            EBS{1,4} = EBS{1,4} + temp{i,2};
        case 'Direct'
            EBS{2,3} = EBS{2,3} + abs(temp{i,2});
            EBS{2,4} = EBS{2,4} + temp{i,2};            
        case 'Distressed Credit'
            EBS{3,3} = EBS{3,3} + abs(temp{i,2});
            EBS{3,4} = EBS{3,4} + temp{i,2};
        case 'Special Situations'
            EBS{4,3} = EBS{4,3} + abs(temp{i,2});
            EBS{4,4} = EBS{4,4} + temp{i,2};
        case 'Stressed/Distressed'
            EBS{5,3} = EBS{5,3} + abs(temp{i,2});
            EBS{5,4} = EBS{5,4} + temp{i,2}; 
        case 'Trading Desk'
            EBS{6,3} = EBS{6,3} + abs(temp{i,2});
            EBS{6,4} = EBS{6,4} + temp{i,2};
    end
end
for i = 1:6
    EBS{i,1} = (EBS{i,3}+EBS{i,4})/2;
    EBS{i,2} = (-EBS{i,3}+EBS{i,4})/2;
end
EBS{7,1} = sum(cell2mat(EBS(1:6,1)));
EBS{7,2} = sum(cell2mat(EBS(1:6,2)));
EBS{7,3} = sum(cell2mat(EBS(1:6,3)));
EBS{7,4} = sum(cell2mat(EBS(1:6,4)));
%  percentage
EBSR = num2cell(cell2mat(EBS)/navn);

%% Exposure by region
temp = sortcell(proc,[5 2]);
EBR = num2cell(zeros(8,4));
for i = 1:size(temp,1)
    switch temp{i,5}
        case 'Asia'
            EBR{1,3} = EBR{1,3} + abs(temp{i,2});
            EBR{1,4} = EBR{1,4} + temp{i,2};
        case 'Canada'
            EBR{2,3} = EBR{2,3} + abs(temp{i,2});
            EBR{2,4} = EBR{2,4} + temp{i,2};            
        case 'European Union'
            EBR{3,3} = EBR{3,3} + abs(temp{i,2});
            EBR{3,4} = EBR{3,4} + temp{i,2};          
        case 'Latin America'
            EBR{4,3} = EBR{4,3} + abs(temp{i,2});
            EBR{4,4} = EBR{4,4} + temp{i,2};
        case 'Non-EU Europe (Ex-UK)'
            EBR{5,3} = EBR{5,3} + abs(temp{i,2});
            EBR{5,4} = EBR{5,4} + temp{i,2};   
        case 'United Kingdom'
            EBR{6,3} = EBR{6,3} + abs(temp{i,2});
            EBR{6,4} = EBR{6,4} + temp{i,2}; 
        case 'United States of America'
            EBR{7,3} = EBR{7,3} + abs(temp{i,2});
            EBR{7,4} = EBR{7,4} + temp{i,2};
    end
end
for i = 1:7
    EBR{i,1} = (EBR{i,3}+EBR{i,4})/2;
    EBR{i,2} = (-EBR{i,3}+EBR{i,4})/2;
end
EBR{8,1} = sum(cell2mat(EBR(1:7,1)));
EBR{8,2} = sum(cell2mat(EBR(1:7,2)));
EBR{8,3} = sum(cell2mat(EBR(1:7,3)));
EBR{8,4} = sum(cell2mat(EBR(1:7,4)));
%  percentage
EBRR = num2cell(cell2mat(EBR)/navn);

%% Sectype
temp = sortcell(raw(2:end,:),[6 2]);
EBST = num2cell(zeros(6,4));
for i = 1:size(temp,1)
    switch temp{i,6}
        case 'Credit Default Swap'
            EBST{1,3} = EBST{1,3} + abs(temp{i,2});
            EBST{1,4} = EBST{1,4} + temp{i,2};
        case 'Common Stock'
            EBST{2,3} = EBST{2,3} + abs(temp{i,2});
            EBST{2,4} = EBST{2,4} + temp{i,2};            
        case 'Bond'
            EBST{3,3} = EBST{3,3} + abs(temp{i,2});
            EBST{3,4} = EBST{3,4} + temp{i,2};
        case 'Bank Debt Loans/Revolvers'
            EBST{3,3} = EBST{3,3} + abs(temp{i,2});
            EBST{3,4} = EBST{3,4} + temp{i,2};
        case 'Futures'
            EBST{4,3} = EBST{4,3} + abs(temp{i,2});
            EBST{4,4} = EBST{4,4} + temp{i,2};            
        case 'Equity Put Option'
            EBST{5,3} = EBST{5,3} + abs(temp{i,2});
            EBST{5,4} = EBST{5,4} + temp{i,2};
        case 'Equity Call Option'
            EBST{5,3} = EBST{5,3} + abs(temp{i,2});
            EBST{5,4} = EBST{5,4} + temp{i,2};
        case 'Preferred Stock'
            EBST{5,3} = EBST{5,3} + abs(temp{i,2});
            EBST{5,4} = EBST{5,4} + temp{i,2}; 
    end
end
for i = 1:5
    EBST{i,1} = (EBST{i,3}+EBST{i,4})/2;
    EBST{i,2} = (-EBST{i,3}+EBST{i,4})/2;
end
EBST{6,1} = sum(cell2mat(EBST(1:5,1)));
EBST{6,2} = sum(cell2mat(EBST(1:5,2)));
EBST{6,3} = sum(cell2mat(EBST(1:5,3)));
EBST{6,4} = sum(cell2mat(EBST(1:5,4)));
%  percentage
EBSTR = num2cell(cell2mat(EBST)/navn);

%% Exposure by Market Cap
temp = sortcell(proc,[7 2]);
EBM = num2cell(zeros(5,4));
for i = 1:size(temp,1)
    switch temp{i,7}
        case '(Sovereign/Govt)Large > 5 Bil'
            EBM{1,3} = EBM{1,3} + abs(temp{i,2});
            EBM{1,4} = EBM{1,4} + temp{i,2};
        case 'LARGE CAP > 5 BILLION'
            EBM{2,3} = EBM{2,3} + abs(temp{i,2});
            EBM{2,4} = EBM{2,4} + temp{i,2};            
        case 'MID CAP 1-5 BILLION'
            EBM{3,3} = EBM{3,3} + abs(temp{i,2});
            EBM{3,4} = EBM{3,4} + temp{i,2};          
        case 'SMALL CAP <1 BILLION'
            EBM{4,3} = EBM{4,3} + abs(temp{i,2});
            EBM{4,4} = EBM{4,4} + temp{i,2};                     
    end
end
for i = 1:4
    EBM{i,1} = (EBM{i,3}+EBM{i,4})/2;
    EBM{i,2} = (-EBM{i,3}+EBM{i,4})/2;
end
EBM{5,1} = sum(cell2mat(EBM(1:4,1)));
EBM{5,2} = sum(cell2mat(EBM(1:4,2)));
EBM{5,3} = sum(cell2mat(EBM(1:4,3)));
EBM{5,4} = sum(cell2mat(EBM(1:4,4)));
%  percentage
EBMR = num2cell(cell2mat(EBM)/navn);

%%  Exposure by country
temp = sortcell(proc,[9 2]);
EBC = num2cell(zeros(23,5));
for i = 1:size(temp,1)
    switch temp{i,9}
        case 'ARGENTINA'
            EBC{1,1} = 'Argentina';
            EBC{1,4} = EBC{1,4} + abs(temp{i,2});
            EBC{1,5} = EBC{1,5} + temp{i,2};
        case 'BERMUDA'
            EBC{2,1} = 'Bermuda';
            EBC{2,4} = EBC{2,4} + abs(temp{i,2});
            EBC{2,5} = EBC{2,5} + temp{i,2};            
        case 'CANADA'
            EBC{3,1} = 'Canada';
            EBC{3,4} = EBC{3,4} + abs(temp{i,2});
            EBC{3,5} = EBC{3,5} + temp{i,2};          
        case 'CROATIA'
            EBC{4,1} = 'Croatia';
            EBC{4,4} = EBC{4,4} + abs(temp{i,2});
            EBC{4,5} = EBC{4,5} + temp{i,2};            
        case 'EUROPE'
            EBC{5,1} = 'Europe';
            EBC{5,4} = EBC{5,4} + abs(temp{i,2});
            EBC{5,5} = EBC{5,5} + temp{i,2};            
        case 'FRANCE'
            EBC{6,1} = 'France';
            EBC{6,4} = EBC{6,4} + abs(temp{i,2});
            EBC{6,5} = EBC{6,5} + temp{i,2};            
        case 'GERMANY'
            EBC{7,1} = 'Germany';
            EBC{7,4} = EBC{7,4} + abs(temp{i,2});
            EBC{7,5} = EBC{7,5} + temp{i,2};            
        case 'GREECE'
            EBC{8,1} = 'Greece';
            EBC{8,4} = EBC{8,4} + abs(temp{i,2});
            EBC{8,5} = EBC{8,5} + temp{i,2};            
        case 'IRELAND'
            EBC{9,1} = 'Ireland';
            EBC{9,4} = EBC{9,4} + abs(temp{i,2});
            EBC{9,5} = EBC{9,5} + temp{i,2};
        case 'JAPAN'
            EBC{10,1} = 'Japan';
            EBC{10,4} = EBC{10,4} + abs(temp{i,2});
            EBC{10,5} = EBC{10,5} + temp{i,2}; 
        case 'NORWAY'
            EBC{11,1} = 'Norway';
            EBC{11,4} = EBC{11,4} + abs(temp{i,2});
            EBC{11,5} = EBC{11,5} + temp{i,2}; 
        case 'PORTUGAL'
            EBC{12,1} = 'Portugal';
            EBC{12,4} = EBC{12,4} + abs(temp{i,2});
            EBC{12,5} = EBC{12,5} + temp{i,2};
        case 'PUERTO RIC'
            EBC{13,1} = 'Puerto Ric';
            EBC{13,4} = EBC{13,4} + abs(temp{i,2});
            EBC{13,5} = EBC{13,5} + temp{i,2}; 
        case 'SCOTLAND'
            EBC{14,1} = 'Scotland';
            EBC{14,4} = EBC{14,4} + abs(temp{i,2});
            EBC{14,5} = EBC{14,5} + temp{i,2}; 
        case 'SLOVAKIA'
            EBC{15,1} = 'Slovakia';
            EBC{15,4} = EBC{15,4} + abs(temp{i,2});
            EBC{15,5} = EBC{15,5} + temp{i,2}; 
        case 'SPAIN'
            EBC{16,1} = 'SPAIN';
            EBC{16,4} = EBC{16,4} + abs(temp{i,2});
            EBC{16,5} = EBC{16,5} + temp{i,2}; 
        case 'SWEDEN'
            EBC{17,1} = 'Sweden';
            EBC{17,4} = EBC{17,4} + abs(temp{i,2});
            EBC{17,5} = EBC{17,5} + temp{i,2}; 
        case 'UAE'
            EBC{18,1} = 'UAE';
            EBC{18,4} = EBC{18,4} + abs(temp{i,2});
            EBC{18,5} = EBC{18,5} + temp{i,2}; 
        case 'UK'
            EBC{19,1} = 'UK';
            EBC{19,4} = EBC{19,4} + abs(temp{i,2});
            EBC{19,5} = EBC{19,5} + temp{i,2}; 
        case 'UKRAINE'
            EBC{20,1} = 'Ukraine';
            EBC{20,4} = EBC{20,4} + abs(temp{i,2});
            EBC{20,5} = EBC{20,5} + temp{i,2}; 
        case 'USA'
            EBC{21,1} = 'USA';
            EBC{21,4} = EBC{21,4} + abs(temp{i,2});
            EBC{21,5} = EBC{21,5} + temp{i,2}; 
        case 'VENEZUELA'
            EBC{22,1} = 'Venezuela';
            EBC{22,4} = EBC{22,4} + abs(temp{i,2});
            EBC{22,5} = EBC{22,5} + temp{i,2}; 
    end
end
for i = 1:22
    EBC{i,2} = (EBC{i,4}+EBC{i,5})/2;
    EBC{i,3} = (-EBC{i,4}+EBC{i,5})/2;
end
EBC{23,1} = 'Total';
EBC{23,2} = sum(cell2mat(EBC(1:22,2)));
EBC{23,3} = sum(cell2mat(EBC(1:22,3)));
EBC{23,4} = sum(cell2mat(EBC(1:22,4)));
EBC{23,5} = sum(cell2mat(EBC(1:22,5)));

%% Calculated record to excel
%xlswrite(file,Time,'Summary','E3:E3');
xlswrite(file,YTD,'Summary','B29:C46');
xlswrite(file,TTL(:,1:3),'Summary','A50:C69');
xlswrite(file,TTS(:,1:3),'Summary','G50:I69');
xlswrite(file,repmat(EBF,[2 1]),'Summary','B73:E74');
xlswrite(file,repmat(EBFR,[2 1]),'Summary','H73:K74');
xlswrite(file,EBI,'Summary','B78:E91');
xlswrite(file,EBIR,'Summary','H78:K91');
xlswrite(file,EBS,'Summary','B95:E101');
xlswrite(file,EBSR,'Summary','H95:K101');
xlswrite(file,EBR,'Summary','B105:E112');
xlswrite(file,EBRR,'Summary','H105:K112');
xlswrite(file,EBST,'Summary','B116:E121');
xlswrite(file,EBSTR,'Summary','H116:K121');
xlswrite(file,EBM,'Summary','B125:E129');
xlswrite(file,EBMR,'Summary','H125:K129');
xlswrite(file,EBC,'Top','A7:E29');
xlswrite(file,TTL,'Top','A33:G52');
xlswrite(file,TTS,'Top','A56:G75');
