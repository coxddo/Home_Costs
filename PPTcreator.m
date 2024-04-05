% Import the PPT API package
import mlreportgen.ppt.*

% Create a presentation container
slides = Presentation('AnalisiCostiCasa', 'myTemplate.pptx');

%% tabella excel iniziale
% Specify the path to your Excel file
excelFilePath = 'excelInput.xlsx'; % replace with your path

% Read the Excel file
dataTable = readtable(excelFilePath);

% Convert the table to a cell array
dataCell = table2cell(dataTable);

% Add column names to the beginning of the cell array
dataCell = [dataTable.Properties.VariableNames; dataCell];

% Create a table for the first slide
tableObj = Table(dataCell);

% Add a slide and add the table to it
slide = add(slides, 'Title and Content');
replace(slide, 'Title', 'Dati di input')
replace(slide, 'Content', tableObj);

%% immagini
% Specify the path to your images
directory= a.folder;
directory = [directory, '\Immagini'];
% directory = [a.folder, '\Immagini']
imageFolderPath = directory; % replace with your path
imageFiles = dir(fullfile(imageFolderPath, '*.svg'));

% Define the position and size for each image
positions = {'0in', '0in', '12cm', '9cm'; '12cm', '0in', '13cm', '9cm'; '0cm', '9cm', '13cm', '9cm'; '12cm', '9cm', '13cm', '9cm'};

% Loop through each image and add it to a slide
k=1;
for i = 1:length(imageFiles)
    % If i is a multiple of 5, create a new slide
    if mod(i, 4) == 1
        slide = add(slides, 'Blank');
        k=1;
    end
    
    % Create a Picture object with the image file
    pic = Picture(fullfile(imageFolderPath, imageFiles(i).name));
    
    % Set the position and size of the picture
    pic.X = positions{mod(k-1, 4)+1, 1};
    pic.Y = positions{mod(k-1, 4)+1, 2};
    pic.Width = positions{mod(k-1, 4)+1, 3};
    pic.Height = positions{mod(k-1, 4)+1, 4};
    
    % Add the picture to the slide
    add(slide, pic);
    k=k+1;
end
%% tabella excel finale
% Specify the path to your Excel file
excelFilePath = 'excelOutput.xlsx'; % replace with your path

% Read the Excel file
dataTable = readtable(excelFilePath);

% Convert the table to a cell array
dataCell = table2cell(dataTable);

% Add column names to the beginning of the cell array
dataCell = [dataTable.Properties.VariableNames; dataCell];

% Create a table for the first slide
tableObj = Table(dataCell);

% Add a slide and add the table to it
slide = add(slides, 'Title and Content');
replace(slide, 'Title', 'Dati output')
replace(slide, 'Content', tableObj);

%% 
% Save and close the presentation
close(slides);
winopen("AnalisiCostiCasa.pptx")