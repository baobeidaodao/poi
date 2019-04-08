package com.baobeidaodao.poi.util;

import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.FileNotFoundException;

/**
 * @author DaoDao
 */
@Slf4j
public class FileUtil {

    public static void checkFile(File file) {
        if (null == file) {
            log.error("file not found!");
            try {
                throw new FileNotFoundException("file not found!");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
                log.error(e.getMessage());
            }
        }
    }
}
