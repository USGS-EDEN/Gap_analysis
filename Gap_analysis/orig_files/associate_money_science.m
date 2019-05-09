%%
% Function to extract account information and distribution of hours from
% BASIS report OM300 from Patti
clear, close all, clc
inDir = 'D:\Administrative\LessOftenUsed\WFP_2017\BasisAnalysis';

monFile = 'FY18 gross project funds.xlsx';
inFileScience = 'tasks_associated_major_science_directions_v3.xlsx';

cd(inDir)

%% Pull in the funding, clear out empties, totals, and plug in all project nums
[monNums,monText] = xlsread(monFile,'funding for WFP');
monText(1:2,:) = [];
bads = [];

for tt = 1:size(monText,1)
    if isempty(monText{tt,1})
        monText{tt,1} = monText{tt-1,1};
    elseif ~isempty(strfind(monText{tt,1},'Total'))
        bads = [bads tt];
    end
end
monText(bads,:) = [];
monNums(bads,:) = [];
clear bads

bads = (isnan(monNums(:,1)) | monNums(:,1)==0) & ...
    (isnan(monNums(:,2)) | monNums(:,2) == 0);
monNums(bads,:) = [];
monText(bads,:) = [];

outCell = cell(size(monText,1),size(monText,2)+size(monNums,2));
for cc = 1:size(monText,1)
    for rr = 1:size(monText,2)
        outCell{cc,rr} = monText{cc,rr};
    end
    for rr = 1:size(monNums,2)
        outCell{cc,rr+size(monText,2)} = monNums(cc,rr);
    end
end
xlswrite('temp.xlsx',outCell)
%MANUALLY ADDED TASKS
%%
[monNums,monText] = xlsread('associate_tasks_to_money.xlsx');
monText(1,:) = [];
monProj = monText(:,1);
monAccount = monText(:,3);
monTask = monNums(:,1);
monPlan = monNums(:,6);
monLate = monNums(:,7);


%% Extract science info
[~,~,raw] = xlsread(inFileScience);
pCodeS = strcmp(raw(1,:),'Project Number');
pCodeS = raw(2:end,pCodeS);

tCodeS = strcmp(raw(1,:),'Task #');
tCodeS = raw(2:end,tCodeS);
tCodeS = cell2mat(tCodeS);
tCodeS(isnan(tCodeS)) = 0;

sciList = raw(1,5:end);
sciPerc = cell2mat(raw(2:end,5:end));
sciPerc(isnan(sciPerc)) = 0;

clear raw

% %%
% sciProjDidntFind = cell(0);
% GXPlan = zeros(1,length(sciList));
% GRPlan = zeros(1,length(sciList));
% GCPlan = zeros(1,length(sciList));
% GXMay18 = zeros(1,length(sciList));
% GRMay18 = zeros(1,length(sciList));
% GCMay18 = zeros(1,length(sciList));
% monPlan(isnan(monPlan)) = 0;
% monLate(isnan(monLate)) = 0;
% monMay18 = zeros(length(pCodeS),length(sciList));
% whichFund = zeros(length(pCodeS),length(sciList));
% 
% for pp = 1:length(pCodeS)
%     it = find(strcmp(pCodeS{pp},monProj) & tCodeS(pp) == monTask);
%     if isempty(it)
%         sciProjDidntFind{length(sciProjDidntFind)+1} = [pCodeS{pp} ', Task ' num2str(tCodeS(pp))];
%         continue
%     end
%     for ii = 1:length(it)
%         TA = monAccount{it(ii)};
%         TA = TA(1:2);
%         eval(['thisPlan = ' TA 'Plan;'])
%         eval(['thisMay18 = ' TA 'May18;'])
%         
%         thisPlan = thisPlan+sciPerc(pp,:).*monPlan(it(ii));
%         thisMay18 = thisMay18+sciPerc(pp,:).*monLate(it(ii));
%         monMay18(pp,:) = monMay18(pp,:)+sciPerc(pp,:).*monLate(it(ii));
%         if strcmp(TA,'GR')
%             whichFund(pp,:) = 1;
%         else
%             whichFund(pp,:) = 2;
%         end
%     end
%     eval([TA 'Plan = thisPlan;'])
%     eval([TA 'May18 = thisMay18;'])
%     clear thisPlan thisMay18
% end

%%
% close all
% figure, subplot(2,1,1)
% b = bar((1:length(sciList)),[GXPlan; GRPlan + GCPlan]'./1000,'stacked');
% set(b(1),'facecolor',[0.2 0.5 0])
% set(b(2),'facecolor','c')
% legend(b,'GX','GR & GC')
% ylim([0 1500])
% title('Sum of Original Gross estimate FY18 B+ Plan')
% ylabel('$1000K')
% 
% subplot(2,1,2)
% b = bar((1:length(sciList)),[GXMay18; GRMay18 + GCMay18]'./1000,'stacked');
% set(b(1),'facecolor',[0.2 0.5 0])
% set(b(2),'facecolor','c')
% legend(b,'GX','GR & GC')
% ylim([0 1500])
% title('Sum of Plan B+ as of 5-23-2018')
% ylabel('$1000K')

monPlan(isnan(monPlan)) = 0;
monLate(isnan(monLate)) = 0;
monProjDidntFind = cell(0);
GXMay18 = zeros(length(monTask),length(sciList));
GRGCMay18 = zeros(length(monTask),length(sciList));
GXPlan = zeros(length(monTask),length(sciList));
GRGCPlan = zeros(length(monTask),length(sciList));

for pp = 1:length(monProj)
    it = find(strcmp(monProj{pp},pCodeS) & tCodeS == monTask(pp));
    if isempty(it)
        monProjDidntFind{length(monProjDidntFind)+1} = [monProj{pp} ', Task ' num2str(monTask(pp))];
        continue
    end
  
    TA = monAccount{pp};
    if strcmp(TA(1:2),'GX')
        GXMay18(pp,:) = monLate(pp)*sciPerc(it,:);
        GXPlan(pp,:) = monPlan(pp)*sciPerc(it,:);
    elseif strcmp(TA(1:2),'GR') || strcmp(TA(1:2),'GC')
        GRGCMay18(pp,:) = monLate(pp)*sciPerc(it,:);
                GRGCPlan(pp,:) = monPlan(pp)*sciPerc(it,:);
    else
        error('Didn''t find code')
    end
end

%%
close all
figure, subplot(2,1,1)
b = bar((1:length(sciList)),[sum(GXPlan,1); sum(GRGCPlan,1)]'./1000,'stacked');
set(b(1),'facecolor',[0.2 0.5 0])
set(b(2),'facecolor','c')
legend(b,'GX','GR & GC')
ylim([0 1500])
title('Sum of Original Gross estimate FY18 B+ Plan')
ylabel('$1000K')

subplot(2,1,2)
b = bar((1:length(sciList)),[sum(GXMay18,1); sum(GRGCMay18,1)]'./1000,'stacked');
set(b(1),'facecolor',[0.2 0.5 0])
set(b(2),'facecolor','c')
legend(b,'GX','GR & GC')
ylim([0 1500])
title('Sum of Plan B+ as of 5-23-2018')
ylabel('$1000K')

