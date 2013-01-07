// Trusted_tool.cpp : Defines the entry point for the console application.
//

#include "del_zip_file.h"

void wtoc(CHAR* Dest, const WCHAR* Source)
{
	int i = 0;

	while(Source[i] != '\0')
	{
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

		if (strlen( extension_type ) > 2){
			return extension_type + 1; 
		}


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
	long len=NULL;
	BYTE *ret=NULL;

	fl= fopen(path, "r+b");
	fseek(fl, 0, SEEK_END);
	len = ftell(fl);
	ret = (BYTE*)malloc(len);

	fseek(fl, 0, SEEK_SET);
	fread(ret, 1, len, fl);
	fclose(fl);
	return ret;
	
}



char is_libre_office(BYTE *hexa_str){
	//50 4B 03 04 14 00 00 08 odt
	const char libre_hex_str[] = {0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x00, 0x06};
	int i=0;
	for(i=0;i<sizeof(libre_hex_str);i++){
		if(hexa_str[i]==libre_hex_str[i]){
			return 0;
		}	
	}
	return 1;
}

char is_MS_office(BYTE *hexa_str){
	//50 4B 03 04 14 00 06 00        XLSX DOCS PPTX
	const char office_hex_str[] = {0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00};
	int i=0;
	for(i=0;i<sizeof(office_hex_str);i++){
		if(hexa_str[i]!=office_hex_str[i]){
			return 0;
		}	
	}
	return 1;
}

char check_magic_number(BYTE *hexa_str){

	char libre=0;
	char office=0;
	libre=is_libre_office(hexa_str);
	office=is_MS_office(hexa_str);
	if(libre && !office )
		return 1;
	else if(office && !libre)
		return 2;
	else
		return 0;
}

char open_xml(char *filepath){
	struct stat file_status;
	char *buf = NULL;
	FILE * pFile=NULL;
	char *ch=NULL;
	char *ch2=NULL;
	char *tmp=NULL;
	char first_pattern[]="<office:scripts>";
	char end_pattern[]="</office:scripts>";
	size_t  nb_first_occr;
	size_t  nb_sec_occr;
	size_t size_rm;

	stat("..\\content.xml", &file_status);
	buf = (char*)malloc(file_status.st_size);
	//buf = (char*)calloc(file_status.st_size,1);
	pFile = fopen ("..\\content.xml","r");
	if (pFile == NULL)  
		printf ("Erreur a l'ouverture du fichier\n");  
	else {  
		fread (buf,1,file_status.st_size,pFile);

		ch = strstr(buf, "<office:scripts>");
		ch2 = strstr(buf, "</office:scripts>");
		nb_first_occr=file_status.st_size-strlen(ch);
		nb_sec_occr=file_status.st_size-strlen(ch2);
		//printf("\n %d \n",nb_first_occr);
		//printf("\n %d \n",nb_sec_occr);
		size_rm=strlen(ch)-strlen(ch2)+strlen(end_pattern);
		tmp=(char*) calloc(file_status.st_size-(size_rm),1);
		
		strncpy(tmp,buf,(size_t)( nb_first_occr+12 ));
		strcat(tmp,ch2+strlen(end_pattern));
		//remove bug
		tmp[file_status.st_size-size_rm]=0;
		printf("\n %s",tmp);
	}
	fclose(pFile);
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



int _tmain(int argc, _TCHAR* argv[])
{
	WIN32_FIND_DATA File;
	HANDLE hSearch;
	int test=0;
	char val=0;
	char office_or_libre=0;
	char tmp[150]={0};
	char * path=NULL;
	BYTE * magic_number=NULL;
	char xml_path[]="content.xml";


	if( argc != 2 )
	{
		_tprintf(TEXT("Usage: %s [target_file]\n"), argv[0]);
		getchar();
		return;
	}

	hSearch = FindFirstFile(argv[1], &File);

	if (hSearch != INVALID_HANDLE_VALUE)
	{
		do {
			//_tprintf(L"%s \n", File.cFileName);

			wtoc(tmp,File.cFileName);
			val = is_macro_file(get_filename_ext(tmp));
			if( val == 1 || val == 2 )
				//printf( "To check" );
				_tprintf(L"%s \n", File.cFileName);
			//DeleteFileFromZIP();
		} while (FindNextFile(hSearch, &File));

		FindClose(hSearch);
	}
	open_xml(xml_path);
	path="test.pptm";
	test = DeleteFileFromZIP(path,"ppt/vbaProject.bin");

	if(test==0)
		printf("\n \n %s", "macro script remove");
	else if(test==1)
		printf("\n \n %s", "nothing remove in archive");
	else
		printf("\n \n %s",  "delete file fail" );

	magic_number=get_magic_number("C:\\Users\\X\\Desktop\\test.pptm");
	office_or_libre = check_magic_number(magic_number);
	modifyFileFromZIP("Destroy.odp", xml_path);
	getchar();
	return 0;


}

