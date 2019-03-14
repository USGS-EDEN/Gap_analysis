% Function to create a people table such that it can be corrected as needed
% BASIS report OM300 from Patti

inDir = 'D:\Administrative\WFP_2017\BasisAnalysis';

%File from August. Chosen because this was the plan and does not include
%some of the movement to FY17 that may not reflect reality in terms of
%distribution.
inFile =  'OM300 Employee assignments 8-15-2017';
sheet = 'FY18 plan 8-15-2017';

cd(inDir)
[~,~,raw] = xlsread(inFile,sheet);

%%
%Extract information

toDo = {'Employee Name';'Employee Title';'Appt Type';'Assigned Hours'};
varName = {'eName';'eTitle';'appt';'hours'};
isRes = cell(size(raw,1)-1,1);
for tt = 1:length(toDo)
    ii = strcmp(raw(1,:),toDo{tt});
    temp = cell(size(raw,1)-1,1);
    for ii2 = 1:size(raw,1)-1
        if any(isnan(raw{ii2+1,ii}))
            temp{ii2,1} = 'Unassigned';
        else
            temp{ii2,1} = raw{ii2+1,ii};
            if strcmp(toDo{tt},'Employee Title')
                if ~isempty(strfind(temp{ii2,1},'Research'))
                    isRes{ii2} = 'Yes';
                else
                    isRes{ii2} = 'No';
                end
            end
        end
    end
    eval([varName{tt} ' = temp;'])
    clear temp
end
hours = cell2mat(hours);
%%
[nameList,ia,ib] = unique(eName);
clear tt ii ii2
allOut = cell(length(nameList)+2,4);
allOut{1,1} = ['Data from ' inFile];
allOut{2,1} = 'Employee Name';
allOut{2,2} = 'Employee Title';
allOut{2,3} = 'Is RGE?';
allOut{2,4} = 'Term/Perm';
allOut{2,5} = 'Total Hours';
allOut(3:end,1) = nameList;
allOut(3:end,2) = eTitle(ia);
allOut(3:end,3) = isRes(ia);
allOut(3:end,4) = appt(ia);
for ii = 1:length(nameList)
    allOut{2+ii,5} = num2str(sum(hours(ib==ii)));
end
%%
xlswrite('employee_list.xlsx',allOut)