clc
clear all
cd 'c:\Users\Prithvish\Desktop\dmd'
allFiles = dir( 'c:\Users\Prithvish\Desktop\dmd' );
allNames = { allFiles.name };
size_file=size(allNames);
for i=3:(size_file(2))
    [pathstr,name,ext] = fileparts(char(allNames(i)))
    [status,sheet]=xlsfinfo(char(allNames(i)));
    mkdir(name);
    cd(name);
    fid_std_dev=fopen(strcat(name,'_std_dev'),'a')
    fid_rms=fopen(strcat(name,'_rms'),'a')
    size_sheet=size(sheet);
    for j=1:size_sheet(2)
        cd ..
        A=xlsread(char(allNames(i)),char(sheet(j)),'');
        cd(name);
        A_size=size(A);
        B=A(8:A_size(1),1:2);
        B_mag=B(:,2);
        B_time=B(:,1);
        % for root mean square 
        root_mean_sq=rms(B_mag);
        fprintf(fid_rms,strcat(char(sheet(j)),' : ',' %f \n'),root_mean_sq);
        % for taking standard deviation of the data
        std_dev=std(B_mag);
        fprintf(fid_std_dev,strcat(char(sheet(j)),' : ',' %f \n'),std_dev);
        % for ploting the data frames
        data_frame=figure('visible','off');
        plot(B_time,B_mag);
        saveas(data_frame,strcat(char(sheet(j)),'_data'),'jpg');
        % for fft of the data
        size_B_mag=size(B_mag)
        fft_data=fft(B_mag,(size_B_mag(1)/3));
        fft_frame=figure('visible','off');
        plot(abs(fft_data));
        saveas(fft_frame,strcat(char(sheet(j)),'_fft'),'jpg');
        % for magnitude of power spectral density
        psd_data_abs=psd(abs(B_mag));
        psd_abs_frame=figure('visible','off');
        plot(psd_data_abs);
        saveas(psd_abs_frame,strcat(char(sheet(j)),'_psd_abs'),'jpg');
        % for angle of power spectral density
        psd_data_angle=psd(angle(B_mag));
        psd_angle_frame=figure('visible','off');
        plot(psd_data_angle);
        saveas(psd_angle_frame,strcat(char(sheet(j)),'_psd_angle'),'jpg');
    end
    fclose(fid_rms);
    fclose(fid_std_dev);
    cd ..
end