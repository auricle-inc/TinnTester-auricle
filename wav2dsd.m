function varargout = wav2dsd(input_file)

if ~exist(input_file,"file")
    error("%s not found\n");
end

[in_path,in_name,in_ext] = fileparts(input_file);
output_file = fullfile(in_path,in_name,'dsf');

% Load the WAV file
[wav_data, wav_sampling_rate] = audioread(input_file);

% Create a dsp.SampleRateConverter object for upsampling
dsd_sampling_rate = 2822400;
src = dsp.SampleRateConverter('InputSampleRate', wav_sampling_rate, 'OutputSampleRate', dsd_sampling_rate);

% Upsample the audio data
upsampled_data = src(wav_data);

% Create a dsp.Quantizer object for 1-bit quantization
quantizer = dsp.Quantizer('Mode', 'Nearest', 'RoundingMethod', 'Floor', 'OverflowAction', 'Wrap', 'NumBits', 1);

% Quantize the upsampled audio data to 1-bit
dsd_data = quantizer(upsampled_data);

% Save as DSD file
audiowrite(output_file, dsd_data, dsd_sampling_rate);

if nargout == 1
    varargout{1} = output_file;
end
