clc
clear all
cd 'c:\Users\Prithvish\Desktop\15th March'
input_file=input('Enter the file name');
    [pathstr,name,ext] = fileparts(char(input_file));
    [status,sheet]=xlsfinfo(char(input_file));
    mkdir(name);
    cd(name);
    fid_std_dev=fopen(strcat(name,'_std_dev'),'a')
    fid_rms=fopen(strcat(name,'_rms'),'a');
    fid_med=fopen(strcat(name,'_med'),'a');
    size_sheet=size(sheet);
    for j=2:size_sheet(2)
        cd ..
        A=xlsread(char(input_file),char(sheet(j)),'');
        A_prev=xlsread(char(input_file),char(sheet(j-1)),'');
        A_1=xlsread(char(input_file),char(sheet(size_sheet(2))),'');
        cd(name);
        A_size=size(A);
        B=A(8:A_size(1),1:2);
        B_prev=A_prev(8:A_size(1),1:2);
        B_1=A_1(8:A_size(1),1:2);
        B_mag=B(:,2);
        B_mag_1=B_1(:,2);
        B_mag_prev=B_prev(:,2)
        B_time=B(:,1);
        % for root mean square 
        root_mean_sq(j)=rms(B_mag);
        fprintf(fid_rms,strcat(char(sheet(j)),' : ',' %f \n'),root_mean_sq(j));
        % for taking standard deviation of the data
        std_dev(j)=std(B_mag);
        fprintf(fid_std_dev,strcat(char(sheet(j)),' : ',' %f \n'),std_dev(j));
        % for ploting the data frames
        data_frame=figure('visible','off');
        plot(B_time,B_mag);
        saveas(data_frame,strcat(char(sheet(j)),'_data'),'jpg');
        % for fft of the data
        size_B_mag=size(B_mag);
        fft_data_1=fft(B_mag_1,2048);
        fft_data_prev=fft(B_mag_prev,2048);
        fft_data=fft(B_mag,2048);
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
        % for the coherence of fft data
        coh_fft=mscohere(fft_data,fft_data_prev);
        coh_fft_frame=figure('visible','off');
        plot(coh_fft);
        saveas(coh_fft_frame,strcat(char(sheet(j)),'_coh_fft'),'jpg');
        % for the coherence of fft data with 1 frame
        [coh_fft_wrt_1,fre]=mscohere(fft_data,fft_data_1,hamming(256),128,256,600);
        coh_fft_wrt_frame_1=figure('visible','off');
        plot(fre,coh_fft_wrt_1);
        saveas(coh_fft_wrt_frame_1,strcat(char(sheet(j)),'_coh_fft_wrt_1'),'jpg');
        % for the differnce in fft data
        diff_fft=abs(fft_data)-abs(fft_data);
        diff_fft_frame=figure('visible','off');
        plot(diff_fft);
        saveas(diff_fft_frame,strcat(char(sheet(j)),'_diff_fft'),'jpg');
        % For  median frequency 
        median_data_fft(j)=median(abs(fft_data(2:25)));
        xx=dist(median_data_fft(j),abs(fft_data)');
        [a,b]=min(xx)
        median_freq(j)=b;
        fprintf(fid_med,strcat(char(sheet(j)),' : ',' %f \n'),median_freq(j));   
    end
    q=1:length(median_freq)
    [r,m,b]=regression(q,median_freq)
    y=m*q+b
    req_plot=figure('visible','on')
    hold on
    plot(q,y)
    plot(q,median_freq)
    saveas(req_plot,'regression_plot','jpg')
    hold off
    rms_fig=figure('visible','off');
    plot(root_mean_sq);
    saveas(rms_fig,'rms_plot','jpg');
    std_fig=figure('visible','off');
    plot(std_dev);
    saveas(std_fig,'std_dev','jpg');
    fclose(fid_rms);
    fclose(fid_std_dev);
    fclose(fid_med);
    cd ..