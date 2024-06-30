function keyPressFcn(src, event, player)
    switch event.Key
        case 'leftarrow'
            volume = max(0, player.UserData.Volume - 0.1);
        case 'rightarrow'
            volume = min(1, player.UserData.Volume + 0.1);
        case 'space'
            disp('Confirmed volume level');
            stop(player);
            return;
        otherwise
            return;
    end
    player.UserData.Volume = volume;
    updateVolume(player, volume);
end

function updateVolume(player, volume)
    player.UserData.SoundWave = volume * player.UserData.OriginalSoundWave;
    play(player);
end

% Usage
figure;
set(gcf, 'KeyPressFcn', @(src, event) keyPressFcn(src, event, player));

% Initial sound setup
fs = 44100;
t = 0:1/fs:1;
soundWave = sin(2 * pi * 1000 * t);
player = audioplayer(soundWave, fs);
player.UserData.Volume = 0.5;
player.UserData.OriginalSoundWave = soundWave;
updateVolume(player, player.UserData.Volume);
