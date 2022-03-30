clear
clc
close all

r_files = dir('CMR*.wav');
r_files = {r_files(:).name};
s_files = dir('CMS*.wav');
s_files = {s_files(:).name};
w_files = dir('CMW*.wav');
w_files = {w_files(:).name};

for idx = 1:length(r_files)
    [data,Fs] = audioread(r_files{idx});
    
    max_mult = 1/max(abs(data));
        data = data.*max_mult;
    
    %audiowrite(data,Fs,16,r_files{idx});
    audiowrite(r_files{idx},data,Fs,'BitsPerSample',16);
end

for idx = 1:length(s_files)
    [data,Fs] = audioread(s_files{idx});
    
    max_mult = 1/max(abs(data));
        data = data.*max_mult;
    
    %audiowrite(data,Fs,16,r_files{idx});
    audiowrite(s_files{idx},data,Fs,'BitsPerSample',16);
end

for idx = 1:length(r_files)
    [data,Fs] = audioread(w_files{idx});
    
    max_mult = 1/max(abs(data));
    data = data.*max_mult;
    
    %audiowrite(data,Fs,16,r_files{idx});
    audiowrite(w_files{idx},data,Fs,'BitsPerSample',16);
end

