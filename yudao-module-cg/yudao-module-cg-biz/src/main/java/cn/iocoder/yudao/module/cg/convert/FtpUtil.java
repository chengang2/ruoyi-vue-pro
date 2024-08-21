package cn.iocoder.yudao.module.cg.convert;

import cn.hutool.extra.ftp.Ftp;
import cn.hutool.extra.ftp.FtpConfig;
import cn.hutool.extra.ftp.FtpMode;
import io.netty.util.CharsetUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPClient;
import org.apache.commons.net.ftp.FTPFile;
import org.apache.commons.net.ftp.FTPReply;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.*;

@Component
@Slf4j
public class FtpUtil {

    @Value("${yudao.ftp.host}")
    private  String host;

    @Value("${yudao.ftp.port}")
    private  int port;

    @Value("${yudao.ftp.username}")
    private  String username;

    @Value("${yudao.ftp.password}")
    private  String password;

    @Value("${yudao.ftp.localpath}")
    public String localpath;
    static String LOCAL_CHARSET = "UTF-8";

    public Ftp initFtp(){
        //打印host的值
        log.info("host:"+host);
        log.info("port:"+port);
        log.info("username:"+username);
        log.info("password:"+password);
        FtpConfig ftpConfig = new FtpConfig();
        ftpConfig.setHost(host);
        ftpConfig.setPort(port);
        ftpConfig.setUser(username);
        ftpConfig.setPassword(password);

        Ftp ftp = new Ftp(ftpConfig,FtpMode.Active);

        return ftp;
    }

    public FTPClient initializeFTPClient() {
        System.out.println("host = " + host);
        System.out.println("port = " + port);
        System.out.println("username = " + username);
        System.out.println("password = " + password);
        FTPClient ftpClient = new FTPClient();
        ftpClient.setAutodetectUTF8(true);
        ftpClient.setCharset(CharsetUtil.UTF_8);
        ftpClient.setControlEncoding(CharsetUtil.UTF_8.name());
        try {
            ftpClient.connect(host, port);
            int replyCode = ftpClient.getReplyCode();
            if (!FTPReply.isPositiveCompletion(replyCode)) {
                System.out.println("Failed to connect to FTP server. Reply code: " + replyCode);
                ftpClient.disconnect();
                return null;
            }
            ftpClient.login(username, password);

            // 设置连接超时时间（毫秒）
            ftpClient.setConnectTimeout(5000);
            // 设置数据传输超时时间（毫秒）
            ftpClient.setDataTimeout(5000);
            // 设置控制保持连接超时时间（毫秒）
            ftpClient.setSoTimeout(5000);

            // 设置为被动模式
            ftpClient.enterLocalPassiveMode();

            ftpClient.setFileType(FTP.BINARY_FILE_TYPE);

            return ftpClient;

        } catch (IOException ex) {
            ex.printStackTrace();
            if (ftpClient.isConnected()) {
                try {
                    ftpClient.disconnect();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            return null;
        }
    }


    public  void listFilesInDirectory(FTPClient ftpClient, String remoteDirectory)  {
        FTPFile[] files = new FTPFile[0];
        try {
            files = ftpClient.listFiles(remoteDirectory);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        if (files != null && files.length > 0) {
            for (FTPFile file : files) {
                if (file.isFile()) {
                    System.out.println("File: " + file.getName());
                } else if (file.isDirectory()) {
                    System.out.println("Directory: " + file.getName());
                }
            }
        } else {
            System.out.println("The directory is empty or does not exist.");
        }
    }

    private static boolean fileExistsOnFTP(FTPClient ftpClient, String filePath) {
        FTPFile[] files = new FTPFile[0];
        try {
            files = ftpClient.listFiles(filePath);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return files.length > 0 && files[0].isFile();
    }

    public boolean uploadFile(FTPClient ftpClient, String localFilePath, String remoteFilePath){
        boolean success = false;
        // 获取远程文件路径的目录部分
        String remoteDirPath = remoteFilePath.substring(0, remoteFilePath.lastIndexOf('/'));

        // 检查并创建远程目录
        if (!createRemoteDirectory(ftpClient, remoteDirPath)) {
            System.out.println("Failed to create remote directory: " + remoteDirPath);
            return false;
        }
        try (InputStream inputStream = new FileInputStream(localFilePath)) {
            success = ftpClient.storeFile(remoteFilePath, inputStream);
            if (success) {
                System.out.println("File has been uploaded successfully.");
            } else {
                System.out.println("Failed to upload the file.");
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
            closeFTPClient(ftpClient);
        }
        return success;
    }
    public boolean createRemoteDirectory(FTPClient ftpClient, String remoteDirPath) {
        try {
            String[] directories = remoteDirPath.split("/");
            StringBuilder path = new StringBuilder();
            for (String dir : directories) {
                if (dir.isEmpty()) continue; // 跳过空目录（根目录）
                path.append("/").append(dir);
                if (!ftpClient.changeWorkingDirectory(path.toString())) {
                    if (!ftpClient.makeDirectory(path.toString())) {
                        return false; // 创建目录失败
                    }
                }
            }
            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }

    public String downloadFile(FTPClient ftpClient, String remoteFilePath, String remotePath,String fileName) {
        String localFilePath = localpath + remotePath;
        File directory = new File(localFilePath);
        if (!directory.exists()) {
            boolean created = directory.mkdirs();
            if (created) {
                log.info("Directory created: " + localpath);
            } else {
                log.error("Failed to create directory: " + localpath);
                return null;
            }
        }
         localFilePath = localFilePath + "/" + fileName;
        try (OutputStream outputStream = new FileOutputStream(localFilePath)) {
            boolean success = ftpClient.retrieveFile(remoteFilePath, outputStream);
            if (success) {
                System.out.println("File has been downloaded successfully.");
                return localFilePath;
            } else {
                System.out.println("Failed to download the file.");
                return null;
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
            closeFTPClient(ftpClient);
        }
        return null;
    }
    public void closeFTPClient(FTPClient ftpClient) {
        if (ftpClient != null && ftpClient.isConnected()) {
            try {
                ftpClient.logout();
                ftpClient.disconnect();
                System.out.println("FTP client has been closed successfully.");
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

//    //下载文件
//    public void downloadFile(Ftp ftp, String remotePath) {
//        ftp.download(remotePath, new File(localpath));
//    }
//
//    //上传文件
//    public boolean uploadFile(Ftp ftp,String remotePath, File file) {
//        boolean uploaded = ftp.upload(remotePath, file);
//        log.info("File uploaded: " + uploaded);
//        return uploaded;
//    }
//
//    public void closeFtp(Ftp ftp){
//        try {
//            ftp.close();
//        } catch (IOException e) {
//            throw new RuntimeException(e);
//        }
//    }

}
