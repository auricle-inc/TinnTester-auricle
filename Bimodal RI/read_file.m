function data = read_file(file_name)


fid = fopen(file_name);
data = fread(fid,9765,'float');
fclose(fid);





