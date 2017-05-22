%Plots and outputs I-V, P-V, R-V and delta-V from the raw data obtained
%from I-V traces measured using a 1000 Ohm load resistor.

DirName = 'FILE_PATH_HERE';
FileName = 'IV-1.dat';

fname = [DirName, FileName];

%import the data and see how many data points there are
M = importdata(fname, '\t');
X = M;
Y = M;
X(:,2) = [];
Y(:,1) = [];
NumRow = size(X, 1);

%manipulate the data to get I, V and Delta (the increment)
i = 1;
for i = 1:NumRow;
    V(i,1) = X(i,1) - Y(i,1); %convert V(app) to the voltage across the sample
    I(i,1) = Y(i,1)/1000; %change the y axis to I from V(res) 
    Delta(i,1) = V(i,1) / 50; %convert V into delta (the drop per superlattice period)
    i = i + 1;
end

%plot I-V
plot(V,I,'-');
title('Current as a function of voltage')
xlabel('Voltage across the sample (V)')
ylabel('Current (I)')

%calculate the power and plot it against the voltage across the sample
P = zeros(NumRow,1);
i = 1;
for i = 1:NumRow
   P(i,1) = V(i,1) * I(i,1);
   i = i +1;
end

figure
plot(V,P,'-')
title('Power as a function of voltage')
xlabel('Voltage across the sample (V)')
ylabel('Power (W)')

%plot power against delta
figure
plot(Delta,P,'-')
title('Power as a function of the Stark Splitting, \Delta')
xlabel('\Delta (V)')
ylabel('Power (W)')

%calculate the resistance and plot it against the voltage across the sample
R = zeros(NumRow,1);
i = 1;
for i = 1:NumRow
   R(i,1) = V(i,1) / I(i,1);
   i = i +1;
end

figure
plot(V,R,'-')
title('Resistance as a function of voltage')
xlabel('Voltage across the sample (V)')
ylabel('Resistance (Ohms)')

%print the four data sets into .dat files in a new directory
OutDirName = [FileName, '_Analysis'];
mkdir(DirName,OutDirName);
Dir = [DirName,OutDirName,'\'];
OutfnameI = [Dir, 'I-V.dat'];
OutfnameP = [Dir, 'Power-V.dat'];
OutfnameR = [Dir, 'Resistance-V.dat'];
OutfnamePDelta = [Dir, 'Power-Increment.dat'];
fidI = fopen(OutfnameI, 'w');
fidP = fopen(OutfnameP, 'w');
fidR = fopen(OutfnameR, 'w');
fidPDelta = fopen(OutfnamePDelta, 'w');
i = 1;
for i = 1:NumRow
    fprintf(fidI, '%d\t', V(i,1));
    fprintf(fidI, '%E\n', I(i,1));
    
    fprintf(fidP, '%d\t', V(i,1));
    fprintf(fidP, '%E\n', P(i,1));
    
    fprintf(fidR, '%d\t', V(i,1));
    fprintf(fidR, '%E\n', R(i,1));
    
    fprintf(fidPDelta, '%d\t', Delta(i,1));
    fprintf(fidPDelta, '%E\n', P(i,1));
    i = i + 1;
end

%label the axes (in easyplot)
fprintf(fidI, '/et    x "Voltage across sample (V)"     ;axis title');
fprintf(fidI, '\n');
fprintf(fidI, '/et    y "Current (I)"        ;axis title');
status = fclose(fidI);

fprintf(fidP, '/et    x "Voltage across sample (V)"     ;axis title');
fprintf(fidP, '\n');
fprintf(fidP, '/et    y "Power (W)"        ;axis title');
status = fclose(fidP);

fprintf(fidR, '/et    x "Voltage across sample (V)"     ;axis title');
fprintf(fidR, '\n');
fprintf(fidR, '/et    y "Resistance (Ohms)"        ;axis title');
status = fclose(fidR);

fprintf(fidPDelta, '/et    x "Stark splitting (V)"     ;axis title');
fprintf(fidPDelta, '\n');
fprintf(fidPDelta, '/et    y "Power (W)"        ;axis title');
status = fclose(fidPDelta);
