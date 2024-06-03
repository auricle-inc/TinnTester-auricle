fs = 48000;
f_start = 100;
f_stop = 20000;

t_dur = 4;

t = 0:1/fs:t_dur;
y = chirp(t,f_start,t_dur,f_stop,'quadratic');

pspectrum(y,fs,'spectrogram','TimeResolution',0.1, ...
    'OverlapPercent',99,'Leakage',0.85)

chirp_filename = fullfile(pwd,sprintf('chirp_fs-%d_fstart-%d_fstop-%d_dur-%d.wav',...
    fs,f_start,f_stop,t_dur));

writeAudioToFile(y,fs,chirp_filename);