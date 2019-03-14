%%
% Function to extract account information and distribution of hours from
% BASIS report OM300 from Patti
clear, close all
inDir = 'D:\Administrative\WFP_2017\BasisAnalysis';

%File from August. Chosen because this was the plan and does not include
%some of the movement to FY17 that may not reflect reality in terms of
%distribution.
inFileHours =  'OM300 Employee assignments 8-15-2017';
sheet = 'FY18 plan 8-15-2017';

inFilePeople = 'employee_list_v2.xlsx';
inFileScience = 'tasks_associated_major_science_directions_v3.xlsx';

cd(inDir)

%%
%Extract information from the OM300
[~,~,raw] = xlsread(inFileHours,sheet);

tCode = strcmp(raw(1,:),'Task Number');
tCode = raw(2:end,tCode);
tCode = cell2mat(tCode);
tCode(isnan(tCode)) = 0;

toDo = {'Employee Name';'Project Number';'Project Title';'Assigned Hours'};
varName = {'eName';'pCode';'pTitle';'hours'};
for tt = 1:length(toDo)
    ii = strcmp(raw(1,:),toDo{tt});
    temp = cell(size(raw,1)-1,1);
    for ii2 = 1:size(raw,1)-1
        if isnan(raw{ii2+1,ii})
            temp{ii2,1} = 'Unassigned';
        else
            temp{ii2,1} = raw{ii2+1,ii};
        end
    end
    eval([varName{tt} ' = temp;'])
    clear temp
end
clear tt ii ii2
hours = cell2mat(hours);
clear raw sheet toDo varName
%%
% Extract information about employees
[~,~,raw] = xlsread(inFilePeople);
eNameT = strcmp(raw(2,:),'Employee Name');
eNameT = raw(3:end,eNameT);
eRGET = strcmp(raw(2,:),'Is RGE?');
eRGET = raw(3:end,eRGET);
ePTT = strcmp(raw(2,:),'Term/Perm');
ePTT = raw(3:end,ePTT);
eRGE = NaN(size(eName,1),1);
eType = NaN(size(eName,1),1);
for ee = 1:length(eNameT)
    iT = strcmp(eName(:,1),eNameT{ee});
    if strcmp(eRGET{ee},'Yes')
        eRGE(iT) = 1;
    elseif strcmp(eRGET{ee}, 'No')
        eRGE(iT) = 0;
    else
        error(['Invalid RGE type for: ' eNameT{ee} ])
    end
    
    if strcmp(ePTT{ee},'FTP')
        eType(iT) = 1;
    elseif strcmp(ePTT{ee}, 'FTO')
        eType(iT) = 0;
    elseif strcmp(ePTT{ee}, 'FEP')
        eType(iT) = -1;
    else
        error(['Invalid Term/Perm type for: ' eNameT{ee} ])
    end
end
clear raw eNameT eRGET ePTT ee

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
%% Generate plots
numProjTasks = zeros(length(sciList),1);
numHoursRGEPerm = zeros(length(sciList),1);
numHoursRGEPermLeave = zeros(length(sciList),1);
numHoursNonRGEPerm = zeros(length(sciList),1);
numHoursNonRGEPermLeave = zeros(length(sciList),1);
numHoursRGETerm = zeros(length(sciList),1);
numHoursRGETermLeave = zeros(length(sciList),1);
numHoursNonRGETerm = zeros(length(sciList),1);
numHoursNonRGETermLeave = zeros(length(sciList),1);
not_done = ones(size(pCode,1),1);
for ss = 1:length(sciList)
    fP = find(sciPerc(:,ss) > 0);
    for tt = 1:length(fP)
        iT = find(strcmp(pCode,pCodeS{fP(tt)}) & (tCode == tCodeS(fP(tt))));
        not_done(iT,1) = 0;
        numProjTasks(ss) = numProjTasks(ss) + 1;
        for pp = 1:length(iT)
            if (strncmp(eName(iT(pp)),'Long',4) || strncmp(eName(iT(pp)),'Stoc',4))
                disp(['Employee left/leaving: ' eName(iT(pp))])
                 numHoursRGEPermLeave(ss) = numHoursRGEPermLeave(ss) + hours(iT(pp))*sciPerc(fP(tt),ss);
            elseif (strncmp(eName(iT(pp)),'Khan',4) || strncmp(eName(iT(pp)),'Lock',4))
                disp(['Employee left/leaving: ' eName(iT(pp))])
                 numHoursRGETermLeave(ss) = numHoursRGETermLeave(ss) + hours(iT(pp))*sciPerc(fP(tt),ss);
            elseif (strncmp(eName(iT(pp)),'Barrer',6)) || strncmp(eName(iT(pp)),'Bren',4)
                numHoursNonRGETermLeave(ss) = numHoursNonRGETermLeave(ss) + hours(iT(pp))*sciPerc(fP(tt),ss);
            elseif (eRGE(iT(pp)) == 1) && (eType(iT(pp)) == 1)
                numHoursRGEPerm(ss) = numHoursRGEPerm(ss) + hours(iT(pp))*sciPerc(fP(tt),ss);
            elseif  (eRGE(iT(pp)) == 0) && (eType(iT(pp)) == 1)
                numHoursNonRGEPerm(ss) = numHoursNonRGEPerm(ss) + hours(iT(pp))*sciPerc(fP(tt),ss);
            elseif (eRGE(iT(pp)) == 1) && (eType(iT(pp)) == 0)
                numHoursRGETerm(ss) = numHoursRGETerm(ss) + hours(iT(pp))*sciPerc(fP(tt),ss);
            elseif  (eRGE(iT(pp)) == 0) && (eType(iT(pp)) == 0)
                numHoursNonRGETerm(ss) = numHoursNonRGETerm(ss) + hours(iT(pp))*sciPerc(fP(tt),ss);
            else
                error(['Invalid employment type for: ' eName{iT(pp)}])
            end
        end
    end
end
didnt_find = find(not_done);
list_nofound = unique(pTitle(didnt_find));
did_find = find(~not_done);
list_found = unique(pTitle(did_find));

%%
close all
!rm project_key_v3.xlsx
figure, subplot(2,1,1)
toDo1 = [numHoursRGEPerm(:) numHoursRGEPermLeave(:) numHoursRGETerm(:) numHoursRGETermLeave(:)];
toDo2 = [numHoursNonRGEPerm(:) numHoursNonRGEPermLeave(:) numHoursNonRGETerm(:) numHoursNonRGETermLeave(:)];
totalHours = sum([toDo1(:); toDo2(:)]);
b = bar(1:length(sciList),100*toDo1./totalHours,'stacked');
set(b(1),'FaceColor','b')
set(b(2),'FaceColor','b','FaceAlpha',0.5)
set(b(3),'FaceColor','k')
set(b(4),'FaceColor','k','FaceAlpha',0.5)
legend(b,'Perm','Perm/Left','Term','Term/Left','location','northeast')
title('RGE')
set(gca,'xtick',1:length(sciList))
ylim([0 12])
ylabel({'% of Total Science'; 'Activity Hours'})
subplot(2,1,2)
b = bar(1:length(sciList),100*toDo2./totalHours,'stacked');
set(b(1),'FaceColor','b')
set(b(2),'FaceColor','b','FaceAlpha',0.5)
set(b(3),'FaceColor','k')
set(b(4),'FaceColor','k','FaceAlpha',0.5)
title('Non-RGE')
ylabel({'% of Total Science'; 'Activity Hours'})
set(gca,'xtick',1:length(sciList))
ylim([0 12])

myOut = cell(length(sciList)+1,2);
myOut{1,1} = 'Number';
myOut{1,2} = 'Major Science Activity';
myOut(2:end,2) = deal(sciList);
for nn = 1:length(sciList)
    myOut{1+nn,1} = [num2str(nn)];
end

figure
text(1,20.5,'Major Science Activities')
text(1,1,myOut(2:end,1))
text(1.5,1,myOut(2:end,2))
axis([0.5 10 -20 20])
box on
axis off
xlswrite('project_key_v3.xlsx',myOut,'ScienceActivities')

myOut = cell(1+length(list_found),1);
myOut{1,1} = 'Included Projects';
myOut(2:1+length(list_found),1) = list_found;
xlswrite('project_key_v3.xlsx',myOut,'Included Projects')

myOut = cell(1+length(list_nofound),1);
myOut{1,1} = 'Excluded Projects';
myOut(2:1+length(list_nofound),1) = list_nofound;
xlswrite('project_key_v3.xlsx',myOut,'Excluded Projects')

% %% By Task
% [allProj,ia,ib] = unique(pCode);
% numTask = NaN(length(allProj),1);
% numPeople = NaN(length(allProj),1);
% for pp = 1:length(allProj)
%     temp = tCode(ib==pp);
%     numTask(pp) = length(unique(temp));
%     numPeople(pp) = length(find(ib==pp));
%     clear temp
% end
%
% allOut = cell(2+sum(numTask),5);
% allOut{1,1} = ['Data from ' inFileHours];
% allOut{2,1} = 'Project Number';
% allOut{2,2} = 'Project Name';
% allOut{2,3} = 'Task Number';
% allOut{2,4} = 'Total People';
% allOut{2,5} = 'Total People-Hours';
%
% cc = 3;
% for pp = 1:length(allProj)
%     taskList = unique(tCode(ib==pp));
%     for tt = 1:length(taskList)
%         iT = find(strcmp(pCode,allProj{pp}) & (tCode == taskList(tt)));
%         allOut{cc,1} = allProj{pp};
%         allOut{cc,2} = pTitle{iT(1)};
%         allOut{cc,3} = num2str(taskList(tt));
%         allOut{cc,4} = num2str(length(iT));
%         allOut{cc,5} = num2str(sum(hours(iT)));
%         cc = cc+1;
%     end
% end
% xlswrite('task_assign_task.xlsx',allOut)
