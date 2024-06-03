function writeAudioToFile(data, samplingRate, fileName)
    % Normalize the data to the range [-1, 1]
    data = data / max(abs(data(:)));
    
    if isrow(data)
        data = data';
    end
    [ch_count,samp_count] = size(data);
    

    % Determine the file encoding type from the filename extension
    [~,~,ext] = fileparts(fileName);
    if strcmpi(ext, '.wav')
        encoding16 = 'int16';
        % encoding24 = 'int24';
    elseif strcmpi(ext, '.flac')
        encoding16 = 'flac';
        % encoding24 = 'flac24';
    else
        error('Unsupported file format. Supported formats: .wav and .flac');
    end

    % Create filenames for 16-bit and 24-bit files
    fileName16 = strrep(fileName, ext, '-16b.wav');
    % fileName24 = strrep(fileName, ext, '-24b.wav');

    % Create an audiowriter object for 16-bit file
    writer16 = dsp.AudioFileWriter(fileName16, 'SampleRate', samplingRate, 'DataType', encoding16);

    % Create an audiowriter object for 24-bit file
    % writer24 = dsp.AudioFileWriter(fileName24, 'SampleRate', samplingRate, 'DataType', encoding24);

    % Write the data to both files
    writer16(int16(2^15*data));
    % writer24(int24(2^23*data));

    % Close the audio writer objects
    release(writer16);
    % release(writer24);
end
