
#ifndef _DEL_ZIP_FILE_
#define _DEL_ZIP_FILE_
#include <windows.h>
#include <tchar.h> 
#include <stdio.h>
#include <stdlib.h>
#include <sys/types.h>
#include <sys/stat.h>
#include <string.h>         /* strrchr */
#include <zlib.h>  
#include <unzip.h>
#include <zip.h>
#include <iowin32.h>
#include <ioapi.h>


#define true 1
#define false 0
char * rename_in_zip(const char * filepath);
char * remove_zip_ext(const char * filepath);
int DeleteFileFromZIP(const char* zip_name, const char* del_file);
int modifyFileFromZIP(const char* zip_name, const char* del_file);


#endif