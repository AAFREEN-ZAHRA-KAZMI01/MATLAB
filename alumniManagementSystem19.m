function alumniManagementSystem()
    % Main GUI setup
    close all;
    global excelFile;
    excelFile = fullfile('C:\Users\AAFREEN ZAHRA KAZMI\Desktop', 'alumniData3.xlsx'); % Path to Excel file
    initializeDatabase(); % Ensure the database exists
    createMainInterface(); % Create the main menu
end

function initializeDatabase()
    % Initialize the database if it doesn't exist
    global excelFile;
    if exist(excelFile, 'file') ~= 2 % Check if the file exists
        headers = {'Name', 'GraduationYear', 'Profession', 'Email', 'PhoneNumber', 'ProfilePicture'}; % Removed ID
        xlswrite(excelFile, headers, 'Sheet1', 'A1'); % Write headers to Excel file
    end
end

function createMainInterface()
    % Create the main interface
    f = figure('Name', 'Alumni Management System', 'NumberTitle', 'off', ...
               'Position', [300, 150, 800, 600], 'MenuBar', 'none', ...
               'Color', [0.8 0.9 1], 'Resize', 'on');

    % Title
    uicontrol('Style', 'text', 'String', 'Alumni Management System', ...
              'FontSize', 16, 'FontWeight', 'bold', ...
              'Position', [200, 500, 400, 40], 'BackgroundColor', [0.8 0.9 1]);

    % Buttons for modules
    uicontrol('Style', 'pushbutton', 'String', '1. Add Alumni Data', ...
              'Position', [250, 400, 200, 50], 'FontSize', 12, ...
              'Callback', @(~, ~) addDataInterface(f));

    uicontrol('Style', 'pushbutton', 'String', '2. Search Alumni by Profession', ...
              'Position', [250, 350, 200, 50], 'FontSize', 12, ...
              'Callback', @(~, ~) searchByProfessionInterface(f));

    uicontrol('Style', 'pushbutton', 'String', '3. Search Alumni by Graduation Year', ...
              'Position', [250, 300, 200, 50], 'FontSize', 12, ...
              'Callback', @(~, ~) searchGraduationYearInterface(f));

    uicontrol('Style', 'pushbutton', 'String', '4. Delete Alumni Data', ...
              'Position', [250, 250, 200, 50], 'FontSize', 12, ...
              'Callback', @(~, ~) deleteDataInterface(f));

    uicontrol('Style', 'pushbutton', 'String', '5. Display Alumni Summary', ...
              'Position', [250, 200, 200, 50], 'FontSize', 12, ...
              'Callback', @(~, ~) displaySummary());

    uicontrol('Style', 'pushbutton', 'String', '6. Export Data to File', ...
              'Position', [250, 150, 200, 50], 'FontSize', 12, ...
              'Callback', @(~, ~) exportData());

    % Close button
    uicontrol('Style', 'pushbutton', 'String', 'Close', ...
              'Position', [250, 50, 200, 50], 'FontSize', 12, ...
              'Callback', @(~, ~) close(f));
end

function addDataInterface(parentFig)
    % Interface for adding alumni data
    if isvalid(parentFig), close(parentFig); end
    f = figure('Name', 'Add Alumni Data', 'NumberTitle', 'off', ...
               'Position', [300, 150, 800, 600], 'Color', [0.9 0.9 1], 'Resize', 'on');

    uicontrol('Style', 'text', 'String', 'Enter Alumni Details:', ...
              'FontSize', 14, 'Position', [250, 520, 200, 30], ...
              'BackgroundColor', [0.9 0.9 1]);

    % Input fields
    fields = {'Name', 'Graduation Year', 'Profession', 'Email', 'Phone Number'};
    yPos = 450:-50:200;
    inputs = cell(size(fields));
    for i = 1:length(fields)
        uicontrol('Style', 'text', 'String', [fields{i}, ':'], ...
                  'Position', [100, yPos(i), 120, 25], 'FontSize', 10, 'BackgroundColor', [0.9 0.9 1]);
        inputs{i} = uicontrol('Style', 'edit', 'Position', [230, yPos(i), 300, 30], 'FontSize', 10);
    end

    % Picture upload section
    uicontrol('Style', 'text', 'String', 'Upload Picture:', ...
              'Position', [550, 350, 120, 25], 'FontSize', 10, 'BackgroundColor', [0.9 0.9 1]);
    uicontrol('Style', 'pushbutton', 'String', 'Upload', ...
              'Position', [550, 300, 100, 30], 'FontSize', 12, ...
              'Callback', @(~, ~) uploadPicture(f));

    % Placeholder for the picture
    ax = axes('Position', [0.75, 0.2, 0.2, 0.4]);

    % Submit button
    uicontrol('Style', 'pushbutton', 'String', 'Submit', ...
              'Position', [350, 50, 100, 40], 'FontSize', 12, ...
              'Callback', @(~, ~) saveUserData(inputs, ax, f));

    % Close button
    uicontrol('Style', 'pushbutton', 'String', 'Close', ...
              'Position', [50, 50, 100, 40], 'FontSize', 12, ...
              'Callback', @(~, ~) close(f));

    % Previous button
    uicontrol('Style', 'pushbutton', 'String', 'Previous', ...
              'Position', [150, 50, 100, 40], 'FontSize', 12, ...
              'Callback', @(~, ~) createMainInterface());
end

function uploadPicture(f)
    [fileName, filePath] = uigetfile({'*.jpg;*.jpeg;*.png', 'Image Files (*.jpg, *.jpeg, *.png)'}); 
    if fileName
        ax = findobj(f, 'Type', 'axes');
        img = imread(fullfile(filePath, fileName));
        imshow(img, 'Parent', ax);
    end
end

function saveUserData(inputs, ax, f)
    % Save user data to Excel
    global excelFile;

    % Read existing data
    [~, ~, existingData] = xlsread(excelFile);

    % Prepare new data
    newData = cell(1, size(existingData, 2)); % Ensure consistent dimensions
    for i = 1:length(inputs)
        newData{i} = inputs{i}.String;
    end

    % Handle profile picture
    img = getimage(ax);
    if ~isempty(img)
        imgPath = fullfile('C:\Users\AAFREEN ZAHRA KAZMI\Desktop', 'profile_picture.jpg');
        imwrite(img, imgPath);
        newData{end} = imgPath;
    else
        newData{end} = ''; % Empty if no picture uploaded
    end

    % Write data to Excel
    existingData = [existingData; newData];
    xlswrite(excelFile, existingData, 'Sheet1');

    % Notify the user and reset interface
    msgbox('Alumni data saved successfully');
    close(f);
    createMainInterface();
end

function searchByProfessionInterface(parentFig)
    % Interface for searching alumni by profession
    if isvalid(parentFig), close(parentFig); end
    f = figure('Name', 'Search by Profession', 'NumberTitle', 'off', ...
               'Position', [300, 150, 800, 600], 'Color', [0.9 0.9 1], 'Resize', 'on');

    uicontrol('Style', 'text', 'String', 'Enter Profession to Search:', ...
              'FontSize', 12, 'Position', [200, 450, 200, 30], 'BackgroundColor', [0.9 0.9 1]);
    professionInput = uicontrol('Style', 'edit', 'Position', [400, 450, 200, 30]);

    uicontrol('Style', 'pushbutton', 'String', 'Search', ...
              'Position', [350, 400, 100, 30], ...
              'Callback', @(~, ~) searchByProfession(professionInput, f)); % Profession is in column 3

    % Close button
    uicontrol('Style', 'pushbutton', 'String', 'Close', ...
              'Position', [50, 50, 100, 40], 'FontSize', 12, ...
              'Callback', @(~, ~) close(f));

    % Previous button
    uicontrol('Style', 'pushbutton', 'String', 'Previous', ...
              'Position', [150, 50, 100, 40], 'FontSize', 12, ...
              'Callback', @(~, ~) createMainInterface());
end

function searchByProfession(professionInput, f)
    % Get the profession entered by the user
    profession = professionInput.String;
    if isempty(profession)
        msgbox('Please enter a profession');
        return;
    end

    % Read the data from Excel
    global excelFile;
    [~, ~, data] = xlsread(excelFile);

    % Find rows with the matching profession
    rows = find(strcmp(data(:, 3), profession));

    % Check if any alumni matched
    if isempty(rows)
        msgbox('No alumni found with this profession');
        return;
    end

    % Display results in a table
    results = data(rows, :);
    uitable('Data', results, 'ColumnName', data(1, :), 'Position', [50, 150, 700, 300]);

    % Notify user
    msgbox(['Found ', num2str(length(rows)), ' alumni(s) with this profession']);
end

function searchGraduationYearInterface(parentFig)
    % Interface for searching alumni by graduation year
    if isvalid(parentFig), close(parentFig); end
    f = figure('Name', 'Search by Graduation Year', 'NumberTitle', 'off', ...
               'Position', [300, 150, 800, 600], 'Color', [0.9 0.9 1], 'Resize', 'on');

    uicontrol('Style', 'text', 'String', 'Enter Graduation Year to Search:', ...
              'FontSize', 12, 'Position', [200, 450, 250, 30], 'BackgroundColor', [0.9 0.9 1]);
    graduationYearInput = uicontrol('Style', 'edit', 'Position', [450, 450, 200, 30]);

    uicontrol('Style', 'pushbutton', 'String', 'Search', ...
              'Position', [350, 400, 100, 30], ...
              'Callback', @(~, ~) searchByGraduationYear(graduationYearInput, f));

    % Close button
    uicontrol('Style', 'pushbutton', 'String', 'Close', ...
              'Position', [50, 50, 100, 40], 'FontSize', 12, ...
              'Callback', @(~, ~) close(f));

    % Previous button
    uicontrol('Style', 'pushbutton', 'String', 'Previous', ...
              'Position', [150, 50, 100, 40], 'FontSize', 12, ...
              'Callback', @(~, ~) createMainInterface());
end


function searchByGraduationYear(gradYearInput, searchFig)
    global excelFile;

    % Get input graduation year
    gradYear = strtrim(get(gradYearInput, 'String'));

    if isempty(gradYear)
        msgbox('Please enter a Graduation Year.', 'Error', 'error');
        return;
    end

    % Read data from the Excel file
    try
        data = readtable(excelFile, 'FileType', 'spreadsheet'); % Removed 'Format' parameter
    catch ME
        if contains(ME.message, 'Unable to open file')
            msgbox('Error: Unable to find or open the Excel file. Ensure the file path and name are correct.', 'Error', 'error');
        else
            msgbox(['Error reading the Excel file: ', ME.message], 'Error', 'error');
        end
        return;
    end

    % Ensure the GraduationYear column exists
    if ~ismember('GraduationYear', data.Properties.VariableNames)
        msgbox('Graduation Year column not found in the data.', 'Error', 'error');
        return;
    end

    % Find rows matching the graduation year
    matchingRows = strcmp(string(data.GraduationYear), gradYear);

    if any(matchingRows)
        % Filter the matching rows
        filteredData = data(matchingRows, :);

        % Display results in a new table
        if isgraphics(searchFig)
            uitable('Parent', searchFig, ...
                    'Data', table2cell(filteredData), ...
                    'ColumnName', filteredData.Properties.VariableNames, ...
                    'Position', [50, 50, 400, 180], ...
                    'RowName', []);
        else
            msgbox('Unable to display data. Search window has been closed.', 'Error', 'error');
        end
    else
        msgbox('No alumni found with this Graduation Year.', 'Information', 'help');
    end
end



function deleteDataInterface(parentFig)
    % Interface for deleting alumni data
    if isvalid(parentFig), close(parentFig); end
    f = figure('Name', 'Delete Alumni Data', 'NumberTitle', 'off', ...
               'Position', [300, 150, 800, 600], 'Color', [0.9 0.9 1], 'Resize', 'on');

    uicontrol('Style', 'text', 'String', 'Enter Year of Alumni to Delete:', ...
              'FontSize', 12, 'Position', [200, 450, 250, 30], 'BackgroundColor', [0.9 0.9 1]);
    nameInput = uicontrol('Style', 'edit', 'Position', [450, 450, 200, 30]);

    uicontrol('Style', 'pushbutton', 'String', 'Delete', ...
              'Position', [350, 400, 100, 30], ...
              'Callback', @(~, ~) deleteData(nameInput)); 

    % Close button
    uicontrol('Style', 'pushbutton', 'String', 'Close', ...
              'Position', [50, 50, 100, 40], 'FontSize', 12, ...
              'Callback', @(~, ~) close(f));

    % Previous button
    uicontrol('Style', 'pushbutton', 'String', 'Previous', ...
              'Position', [150, 50, 100, 40], 'FontSize', 12, ...
              'Callback', @(~, ~) createMainInterface());
end

function deleteData(nameInput, f)
    % Delete alumni data
    global excelFile;
    name = nameInput.String;
    [~, ~, data] = xlsread(excelFile);
    
    % Find matching entry and delete it
    matchIndex = [];
    for i = 2:size(data, 1)
        if strcmpi(data{i, 1}, name)
            matchIndex = [matchIndex, i];
        end
    end
 
    if isempty(matchIndex)
        msgbox('No matching alumni found to delete.');
    else
        data(matchIndex, :) = []; % Remove matching rows
        xlswrite(excelFile, data, 'Sheet1'); % Write back the updated data
        msgbox('Alumni data deleted successfully.');
    end
end





function displaySummary()
    % Display the summary of alumni data
    global excelFile;
    [~, ~, data] = xlsread(excelFile);

    % Calculate total number of alumni
    totalAlumni = size(data, 1) - 1; % Exclude header row
    professions = unique(data(2:end, 3)); % Unique professions

    % Display total number of alumni
    msgbox(['Total number of alumni: ', num2str(totalAlumni)]);

    % Create a bar graph for profession distribution
    professionCounts = zeros(size(professions));
    for i = 1:length(professions)
        professionCounts(i) = sum(strcmp(data(:, 3), professions{i}));
    end

    figure;
    bar(professionCounts);
    set(gca, 'XTickLabel', professions);
    title('Alumni Profession Distribution');
    xlabel('Profession');
    ylabel('Count');
end

function exportData()
    % Export the alumni data to a text file
    global excelFile;
    
    % Read the data from Excel
    [~, ~, data] = xlsread(excelFile);

    if isempty(data)
        msgbox('No alumni data available');
    else
        % Open a file dialog to choose the save location
        [fileName, filePath] = uiputfile('*.txt', 'Save Data As');
        
        if fileName
            try
                % Open the text file for writing
                fileID = fopen(fullfile(filePath, fileName), 'w');
                
                % Write headers (column names)
                headers = {'Name', 'GraduationYear', 'Profession', 'Email', 'PhoneNumber', 'ProfilePicture'};
                fprintf(fileID, '%s\t', headers{:});
                fprintf(fileID, '\n');
                
                % Write the data
                for i = 1:size(data, 1)
                    fprintf(fileID, '%s\t', data{i, :});
                    fprintf(fileID, '\n');
                end
                
                % Close the file
                fclose(fileID);
                
                % Show success message
                msgbox('Data exported successfully');
            catch ME
                msgbox(['Error exporting data: ', ME.message]);
            end
        end
    end
end

