%% MARKET BETA CALCULATION FROM RETURNS OF POSITIONS IN CSSM FUND, CYRUS CAPITAL PARTNERS
%  AUTHOR: YORK YU
%  DATE: 07/22/2014
%  ANALYSIS BACKGROUND INFORMATION UP TO 07/21/2014;
%  SUPERVISOR: ROB JACKSON
%  INTERNAL USE ONLY
%  MODEL: APT WITH PANAL DATA AND DDUMMY VARIABLES
%  BENCHMARK: SP500 

%% LOAD IN DATA
%RETURN = xlsread('C:\Users\yyu\Desktop\MARKET BETA\CMF2 record','historical return');
RETURN = xlsread('C:\Users\yyu\Desktop\MARKET BETA\CMF2 record','historical return 2yrs');
%RETURN = xlsread('C:\Users\yyu\Desktop\MARKET BETA\CSSM record','historical return');
%RETURN = xlsread('C:\Users\yyu\Desktop\MARKET BETA\CSSM record','historical return 2yrs');
TIME_PANEL_DATA = xlsread('C:\Users\yyu\Desktop\MARKET BETA\time panel data 2y');
%TIME_PANEL_DATA = xlsread('C:\Users\yyu\Desktop\MARKET BETA\time panel data 6m');

%% Factors for regression
% panel data that change over time
vix = TIME_PANEL_DATA(:,1);
volume = TIME_PANEL_DATA(:,2);
trend = TIME_PANEL_DATA(:,3);
smb = TIME_PANEL_DATA(:,4);
hml = TIME_PANEL_DATA(:,5);
treasury = TIME_PANEL_DATA(:,6);
oil = TIME_PANEL_DATA(:,7);
iron_ore = TIME_PANEL_DATA(:,8);

% storing beta, both upside and downside, residual_squared, and numbers of securities
nos = length(RETURN(1,:))-1;
beta = zeros(nos,10);
residualsqrt = zeros(nos,1);

% benchmark sp500 return
sp500 = RETURN(:,end);
for i = 1:nos
    r_i = RETURN(:,i); % return on security i
    time = sum(~isnan(r_i));
    if time <= 8 % singularity
        beta(i,1) = 0;
        beta(i,2) = 0;
        continue
    end
    r = r_i(1:time,1)*100;
    sp500_i = sp500(end-time+1:end,1);
    upside_r = (sp500_i+abs(sp500_i))/2*100;
    downside_r = (sp500_i-abs(sp500_i))/2*100;
    factor = [ones(time,1) upside_r downside_r vix(end-time+1:end,1) ...
        volume(end-time+1:end,1) smb(end-time+1:end,1) hml(end-time+1:end,1)...
        treasury(end-time+1:end,1) oil(end-time+1:end,1) iron_ore(end-time+1:end,1)];
    %factor = [ones(time,1) sp500_i];
    [b,bint,res,rint,stats] = regress(r,factor);
    beta(i,:) = b;
    residualsqrt(i) = stats(1);
end

corr_factor = [vix volume smb hml treasury];
corr_m = zeros(1,5);
for j = 1:5
    corr_m(j) = regress(corr_factor(:,j),sp500);
end

final_beta = (beta(:,2)+beta(:,3))/2;
