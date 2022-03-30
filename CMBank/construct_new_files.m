load('calibration_data_cm.mat');

atten_value_s = calib_data_cm(1,:);
atten_value_r = calib_data_cm(2,:);
atten_value_w = calib_data_cm(3,:);

%% Get File Names
r_files = dir('CMR*.wav');
r_files = {r_files(:).name};
s_files = dir('CMS*.wav');
s_files = {s_files(:).name};
w_files = dir('CMW*.wav');
w_files = {w_files(:).name};

%% Get and calculate attenuation values
scale = 95;
atten_value_r = scale - atten_value_r;
%atten_value_r(atten_value_r >= 0) = 0;
atten_value_r = 10.^(atten_value_r./20);

%% Apply to each
for idx = 1:length(r_files)

[data,Fs] = wavread(r_files{idx});

max_mult = 1/max(abs(data));
if (max_mult >= atten_value_r(idx)) || atten_value_r(idx) <= 1
    data = data.*atten_value_r(idx);
else
    data = data.*max_mult;
end

wavwrite(data,Fs,16,r_files{idx});

end

%atten_value_s = [];
atten_value_s = scale - atten_value_s;
%atten_value_s(atten_value_s >= 0) = 0;
atten_value_s = 10.^(atten_value_s./20);


for idx = 1:length(s_files)

[data,Fs] = wavread(s_files{idx});

max_mult = 1/max(abs(data));

if max_mult >= atten_value_s(idx) || atten_value_s(idx) <= 1
    data = data.*atten_value_s(idx);
else
    data = data.*max_mult;
end

wavwrite(data,Fs,16,s_files{idx});

end

%atten_value_w = [];
atten_value_w = scale - atten_value_w;
%atten_value_w(atten_value_w >= 0) = 0;
atten_value_w = 10.^(atten_value_w./20);


for idx = 1:length(w_files)

[data,Fs] = wavread(w_files{idx});
max_mult = 1/max(abs(data));

if (max_mult >= atten_value_w(idx)) || atten_value_w(idx) <= 1 
    data = data.*atten_value_w(idx);
else
    data = data.*max_mult;
end

wavwrite(data,Fs,16,w_files{idx});

end














