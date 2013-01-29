// Trusted_tool.cpp : Defines the entry point for the console application.
//

#include "del_zip_file.h"
#include <stdlib.h>
#include <crtdbg.h>

#define BUFSIZE 4096
#define _CRTDBG_MAP_ALLOC

void wtoc(CHAR* Dest, const WCHAR* Source)
{
	int i = 0;

	while(Source[i] != '\0'){
		Dest[i] = (CHAR)Source[i];
		++i;
	}
	Dest[i] = '\0';
}

const char *get_filename_ext(const char *filename) {
	const char *extension_type = strrchr(filename, '.');
	if (!extension_type) {
		/* no extension */
		return "10";
	} else {

		if (strlen( extension_type ) > 2)
			return extension_type + 1; 

	}
	return extension_type;
}

char is_macro_file(const char *extension_type){

	if(strcmp(extension_type,"docm") == 0 || 
		strcmp(extension_type,"xlsm") == 0 ||
		strcmp(extension_type,"pptm") == 0 ){
			return 1;
	}
	else if(strcmp(extension_type,"odt") == 0 || 
		strcmp(extension_type,"odp") == 0 ||
		strcmp(extension_type,"ods") == 0 ){
			return 1;
	}
	else{
		return 0;
	}
}


BYTE * get_magic_number(const char *path){

	FILE *fl=NULL;
	long len=0;
	BYTE *ret=NULL;

	fl= fopen(path, "r+b");
	if(fl==NULL) //error file no exist
		return 0;
	if(fseek(fl, 0, SEEK_END))
		return 0;
	len = ftell(fl);
	fseek(fl, 0, SEEK_SET);
	ret = (BYTE*)malloc(sizeof(BYTE)*len);
	if(ret == NULL) 
		return 0;
	
	if (fread(ret, 1, len, fl) != len) 
		return 0;
	fclose(fl);
	return ret;
}



char is_libre_office(BYTE *hexa_str){
	//50 4B 03 04 14 00 00 08 odt  //byte 0x00=0x08
	const BYTE libre_hex_str[] = {0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x00, 0x08, 0x00};
	int i=0;
	for(i=0;i<sizeof(libre_hex_str);i++){
		if(hexa_str[i]!=libre_hex_str[i]){
			return 0;
		}	
	}
	return 1;
}

char is_MS_office(BYTE *hexa_str){
	//50 4B 03 04 14 00 06 00 08     XLSX DOCS PPTX
	//50 4B 03 04 14 00 00 00 08     XLSM
 const BYTE office_hex_str[] = {0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00, 0x08}; //magic number
 const BYTE office_hex_str2[] = {0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00, 0x08};
	int i=0;
	int ret=0;

	for(i=0;i<sizeof(office_hex_str);i++){
		if(hexa_str[i]!=office_hex_str[i]){
			ret++;
			break;
		}	
	}
	for(i=0;i<sizeof(office_hex_str);i++){
		if(hexa_str[i]!=office_hex_str2[i]){
			ret++;
			break ;
		}	
	}
	return ret; // if ret=2 no MS office file
	
}

char is_old_MS_office(BYTE *hexa_str){
	//D0 CF 11 E0 A1 B1 1A E1 00 00
 const BYTE office_hex_str[] = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}; //magic number
	int i=0;
	for(i=0;i<sizeof(office_hex_str);i++){
		if(hexa_str[i]!=office_hex_str[i] ){
			return 1;
		}	
	}
	return -1;
}


char check_magic_number(BYTE *hexa_str){

	char libre=0;
	char office=0;
	char old_MS=0;
	libre=is_libre_office(hexa_str);
	office=is_MS_office(hexa_str);
	old_MS=is_old_MS_office(hexa_str);
	free(hexa_str);
	if(old_MS==-1)
		return -1;
	else if(libre && office==2 )
		return 1; //libre office
	else if(office==1 && !libre)
		return 2; //MS office
	else
		return 0;
}



long fsize(FILE * file)
{
	long size;
	if(fseek(file, 0, SEEK_END))
		return -1;

	size = ftell(file);
	if(size < 0)
		return -1;

	if(fseek(file, 0, SEEK_SET))
		return -1;

	return size;
}


/*argv[1] folder to scan*/
int _tmain(int argc, _TCHAR* argv[])
{
	WIN32_FIND_DATA File;
	HANDLE hSearch;
	char office_or_libre=0;
	char tmp[150]={0};
	BYTE * magic_number={0};


	TCHAR  buffer[BUFSIZE]=TEXT(""); 
	CHAR  path[BUFSIZE]={0};
	TCHAR  buf2[BUFSIZE]={0};
	TCHAR** lppPart={NULL};


	if( argc != 2 )	{
		_tprintf(TEXT("Usage: %s [target_file]\n"), argv[0]);
		getchar();
		return;
	}
	


	if ( GetFullPathName(argv[1], BUFSIZE,buffer,  lppPart) == 0)  {
		// Handle an error condition.
		printf ("GetFullPathName failed (%d)\n", GetLastError());
		return;
	}
	else 
		_tprintf(TEXT("The full path name is:  %s\n"), buffer);



	_tcscat(buf2,argv[1]);
	_tcscat(buf2,L"\\*.*"); //search all file in folder

	hSearch = FindFirstFile(buf2, &File);

	if (hSearch != INVALID_HANDLE_VALUE)
	{
		do {
			
			if(File.dwFileAttributes==FILE_ATTRIBUTE_DIRECTORY){
				//sub folder
			}	
			//_tprintf(L"%s \n", File.cFileName);
			wcstombs( path, argv[1], BUFSIZE );
			wtoc(tmp,File.cFileName);
			strcat(path,tmp);
			if((magic_number = get_magic_number(path))!=0){
				office_or_libre = check_magic_number(magic_number);

				if( office_or_libre == 1  ){ //libre office
					_tprintf(L"%s \tis Libre Office document\n", File.cFileName);
					modifyFileFromZIP(path,"content.xml");
					//_tprintf(L"%s is clear macro remove\n", File.cFileName);
				}
				else if( office_or_libre == 2  ) {// MS office

					_tprintf(L"%s is Microsoft Office document\n", File.cFileName);
					DeleteFileFromZIP(path,"xl/vbaProject.bin");	//try to remove vba file in word, excel and pwerpoint
					DeleteFileFromZIP(path,"word/vbaProject.bin");
					DeleteFileFromZIP(path,"ppt/vbaProject.bin");   //remove vba file
					modifyFileFromZIP(path,"ppt/_rels/presentation.xml.rels"); //fix bad opening by removing binding with vba file
					 
					/*if(test==0)
					printf("\n %s", "macro script remove");
					else if(test==1)
					printf("\n %s", "nothing remove in archive");
					else
					printf("\n %s",  "delete file fail" );*/
				}
				else if( office_or_libre == -1  ) {

					_tprintf(L"%s is Old Microsoft Office document\n", File.cFileName);
					remove(path);
					_tprintf(L"%s is removed\n", File.cFileName);
				}
				else
					_tprintf(L"%s is no a Libre or Microsoft Office\n", File.cFileName);
			}
			else {
				//_tprintf(L"%s is no a good path\n", File.cFileName);
			}

		} while (FindNextFile(hSearch, &File));

		FindClose(hSearch);
	}


	getchar();
	return 0;
}

