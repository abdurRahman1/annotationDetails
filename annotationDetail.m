%%
%Get content of the directory
contents = dir ('*.jpg');
excelFilename = 'test.xlsx';

%Preallocating Variables
height = zeros(numel(contents),1);
width =  zeros(numel(contents),1);
nameFile{numel(contents),1} = [];
annPoint = zeros(numel(contents),1);

%Loop to iterate through images
for i = 1:numel(contents)
    
    %Processing for Image details
    filename = contents(i).name;
    readImage1 = imread(filename);
    [path name] = fileparts(filename);
    out_filename = sprintf('%s.jpg', name);
    [x1,y1,z1] = size(readImage1);
    height(i) = x1;
    width(i) = y1;
    nameFile(i) = {out_filename};
    
    %Processing for Annotation details
    matFileName = sprintf('%s_ann.mat', name);
    m = matfile(matFileName);
    [numrow, numcol] = size(m, 'annPoints');
    annPoint(i) = numrow;
end

%Writing Values to Excel file
xlswrite(excelFilename,nameFile,1,'A2');
xlswrite(excelFilename,height,1,'B2');
xlswrite(excelFilename,width,1,'C2');
xlswrite(excelFilename,annPoint,1,'D2');
    
    cell4height=[ 'B' num2str(i)];
    cell4width = [ 'C' num2str(i)];
    cell4name = [ 'A' num2str(i)];
    A = [out_filename];
    
    xlswrite(excelFilename,{out_filename},1,cell4name);
    xlswrite(excelFilename,x1,1,cell4height);
    xlswrite(excelFilename,y1,1,cell4width);

m = matfile('Image002_ann_Copy.mat');
[numrow, numcol] = size(m, 'annPoints');

filename = 'Image002_ann_Copy.mat';
m = matfile(filename);




filename = 'test.xlsx';
readImage1=imread('Image044.jpg');

[x1,y1,z1] = size(readImage1);


xlswrite(filename,x1)