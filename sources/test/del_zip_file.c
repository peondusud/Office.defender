
///////////////////////////////////
//  Ivan A. Krestinin            //
//  Crown_s Soft                 //
//  http://www.crown-s-soft.com  //
///////////////////////////////////


#include "del_zip_file.h"



int modify_xml(void *file,int original_size){
	char *ch=NULL;
	char *ch2=NULL;
	char *ch3=NULL;
	int cheat_offset;
	char first_pattern[]="<office:scripts>";
	char end_pattern[]="</office:scripts>";
	char ppt_binding2remove[]="<Relationship Id=\"rId8\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/>";

	cheat_offset=strlen((char*)file)-original_size;
	ch = strstr((char*)file, "<office:scripts>");
	ch2 = strstr((char*)file, "</office:scripts>");
	ch3 = strstr((char*)file,ppt_binding2remove);

	//if maccro
	if(ch!=NULL && ch2!=NULL)
	memmove_s(ch,	strlen(ch2)-strlen(end_pattern)+1,		ch2+strlen(end_pattern),			strlen(ch2)-strlen(end_pattern)+1);

	//remove vba binding
	if(ch3!=NULL )
	memmove_s(ch3, strlen(ch3)-strlen(ppt_binding2remove)+1,	ch3+strlen(ppt_binding2remove),		strlen(ch3)-strlen(ppt_binding2remove)+1);

	return (strlen((char*)file)-cheat_offset);
}



int DeleteFileFromZIP(const char* zip_name, const char* del_file)
{
	BOOL some_was_del = false;
	char* glob_comment = NULL;
	char* tmp_name ;
	zipFile szip;
	zipFile dzip;
	zlib_filefunc_def ffunc;
	zip_fileinfo zfi;
	unz_file_info unzfi;
	unz_global_info glob_info;

	int n_files,rv,method,level,size_local_extra,sz;
	char* extrafield;
	char* commentary;

	char dos_fn[MAX_PATH];
	char fn[MAX_PATH];
	void* local_extra;
	void* buf;


	// change name for temp file
	tmp_name = (char*)malloc((strlen(zip_name) + 5)*sizeof(char));
	strcpy(tmp_name, zip_name);
	//strcpy_s(tmp_name,strlen(zip_name), zip_name);
	strncat(tmp_name, ".tmp", 5);

	// open source and destination file
	fill_win32_filefunc(&ffunc);

	szip = unzOpen(zip_name);
	if (szip==NULL) { 
		free(tmp_name);
		return -1;
	}

	dzip = zipOpen(tmp_name, APPEND_STATUS_CREATE);
	if (dzip==NULL) {
		unzClose(szip);
		free(tmp_name);
		return -1;
	}

	// get global commentary

	if (unzGetGlobalInfo(szip, &glob_info) != UNZ_OK) { 
		zipClose(dzip, NULL);
		unzClose(szip);
		free(tmp_name);
		return -1;
	}


	if (glob_info.size_comment > 0)
	{
		glob_comment = (char*)malloc(glob_info.size_comment+1);
		if ((glob_comment==NULL)&&(glob_info.size_comment!=0)) {
			zipClose(dzip, NULL);
			unzClose(szip);
			free(tmp_name);
			return -1;
		}

		if ((unsigned int)unzGetGlobalComment(szip, glob_comment, glob_info.size_comment+1) != glob_info.size_comment)  { 
			zipClose(dzip, NULL);
			unzClose(szip); 
			free(glob_comment);
			free(tmp_name);
			return -1;
		}
	}

	// copying files
	n_files = 0;

	rv = unzGoToFirstFile(szip);
	while (rv == UNZ_OK)
	{
		// get zipped file info

		if (unzGetCurrentFileInfo(szip, &unzfi, dos_fn, MAX_PATH, NULL, 0, NULL, 0) != UNZ_OK) break;

		OemToCharA(dos_fn, fn);

		// if not need delete this file
		if (_stricmp(fn, del_file)==0) // lowercase comparison
			some_was_del = true;
		else
		{
			extrafield = (char*)malloc(unzfi.size_file_extra);
			if ((extrafield==NULL)&&(unzfi.size_file_extra!=0))
				break;
			commentary = (char*)malloc(unzfi.size_file_comment);

			if ((commentary==NULL)&&(unzfi.size_file_comment!=0)) {
				free(extrafield);
				break;}

			if ( unzGetCurrentFileInfo(szip, &unzfi, dos_fn, MAX_PATH, extrafield, unzfi.size_file_extra, commentary, unzfi.size_file_comment) != UNZ_OK) {
				free(extrafield);
				free(commentary);
				break;}

			// open file for RAW reading
			if (unzOpenCurrentFile2(szip, &method, &level, 1)!=UNZ_OK) {
				free(extrafield);
				free(commentary);
				break;
			}

			size_local_extra = unzGetLocalExtrafield(szip, NULL, 0);
			if (size_local_extra<0) {
				free(extrafield);
				free(commentary);
				break;
			}

			local_extra = malloc(size_local_extra);
			if ((local_extra==NULL)&&(size_local_extra!=0)) {
				free(extrafield);
				free(commentary);
				break;}
			if (unzGetLocalExtrafield(szip, local_extra, size_local_extra)<0) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				break;}

			// this malloc may fail if file very large
			buf = malloc(unzfi.compressed_size);
			if ((buf==NULL)&&(unzfi.compressed_size!=0)) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				break;}

			// read file
			sz = unzReadCurrentFile(szip, buf, unzfi.compressed_size);
			if ((unsigned int)sz != unzfi.compressed_size) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				free(buf);
				break;}

			// open destination file
			//dest , src ,size
			memcpy(&zfi.tmz_date, &unzfi.tmu_date, sizeof(tm_unz));
			zfi.dosDate = unzfi.dosDate;
			zfi.internal_fa = unzfi.internal_fa;
			zfi.external_fa = unzfi.external_fa;

			if (zipOpenNewFileInZip2 (dzip, dos_fn, &zfi, local_extra, size_local_extra, extrafield, unzfi.size_file_extra, commentary, method, level, 1)!=UNZ_OK) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				free(buf);
				break;}

			// write file
			if (zipWriteInFileInZip(dzip, buf, unzfi.compressed_size)!=UNZ_OK) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				free(buf);
				break;}

			if (zipCloseFileInZipRaw(dzip, unzfi.uncompressed_size, unzfi.crc)!=UNZ_OK) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				free(buf);
				break;}

			if (unzCloseCurrentFile(szip)==UNZ_CRCERROR) {
				free(extrafield);
				free(commentary);
				free(local_extra); 
				free(buf);
				break;}

			free(commentary);
			free(buf);
			free(extrafield);
			free(local_extra);

			n_files ++;
		}

		rv = unzGoToNextFile(szip);
	}

	zipClose(dzip, glob_comment);
	unzClose(szip);

	free(glob_comment);


	// if fail
	if ( (!some_was_del) || (rv!=UNZ_END_OF_LIST_OF_FILE) )
	{
		remove(tmp_name);		
		free(tmp_name);
		return 1;
	}

	remove(zip_name);
	if (rename(tmp_name, zip_name) != 0)
	{
		free(tmp_name);
		return -1;
	}

	// if all files were deleted
	if (n_files==0)
		remove(zip_name);

	free(tmp_name);
	return 0;
}


int modifyFileFromZIP(const char* zip_name, const char* del_file)
{

	BOOL some_was_del = false;
	char* glob_comment = NULL;
	char* tmp_name ;
	zipFile szip;
	zipFile dzip;
	zlib_filefunc_def ffunc;
	zip_fileinfo zfi;
	unz_file_info unzfi;
	unz_global_info glob_info;

	int n_files,rv,method,level,size_local_extra,sz;
	char* extrafield;
	char* commentary;
	char dos_fn[MAX_PATH];
	char fn[MAX_PATH];
	void* local_extra;
	void* buf;


	
	// change name for temp file
	tmp_name = (char*)malloc((strlen(zip_name) + 5)*sizeof(char));
	strcpy(tmp_name, zip_name);
	strncat(tmp_name, ".tmp", 5);

	// open source and destination file
	fill_win32_filefunc(&ffunc);

	szip = unzOpen(zip_name);
	if (szip==NULL) { 
		free(tmp_name);
		return -1;
	}

	dzip = zipOpen(tmp_name, APPEND_STATUS_CREATE);
	if (dzip==NULL) {
		unzClose(szip);
		free(tmp_name);
		return -1;
	}

	// get global commentary

	if (unzGetGlobalInfo(szip, &glob_info) != UNZ_OK) { 
		zipClose(dzip, NULL);
		unzClose(szip);
		free(tmp_name);
		return -1;
	}


	if (glob_info.size_comment > 0)
	{
		glob_comment = (char*)malloc(glob_info.size_comment+1);
		if ((glob_comment==NULL)&&(glob_info.size_comment!=0)) {
			zipClose(dzip, NULL);
			unzClose(szip);
			free(tmp_name);
			return -1;
		}

		if ((unsigned int)unzGetGlobalComment(szip, glob_comment, glob_info.size_comment+1) != glob_info.size_comment)  { 
			zipClose(dzip, NULL);
			unzClose(szip); 
			free(glob_comment);
			free(tmp_name);
			return -1;
		}
	}

	// copying files
	n_files = 0;

	rv = unzGoToFirstFile(szip);
	while (rv == UNZ_OK)
	{
		// get zipped file info

		if (unzGetCurrentFileInfo(szip, &unzfi, dos_fn, MAX_PATH, NULL, 0, NULL, 0) != UNZ_OK) break;

		OemToCharA(dos_fn, fn);

		// if not need delete this file
		if (_stricmp(fn, del_file)==0) // lowercase comparison
		{
			extrafield = (char*)malloc(unzfi.size_file_extra);
			if ((extrafield==NULL)&&(unzfi.size_file_extra!=0))
				break;

			commentary = (char*)malloc(unzfi.size_file_comment);
			if ((commentary==NULL)&&(unzfi.size_file_comment!=0)) {
				free(extrafield);
				break;}

			if ( unzGetCurrentFileInfo(szip, &unzfi, dos_fn, MAX_PATH, extrafield, unzfi.size_file_extra, commentary, unzfi.size_file_comment) != UNZ_OK) {
				free(extrafield);
				free(commentary);
				break;}

			// open file for RAW reading
			if (unzOpenCurrentFile2 (szip, &method, &level, 0)!=UNZ_OK) {
				free(extrafield);
				free(commentary);
				break;
			}

			size_local_extra = unzGetLocalExtrafield(szip, NULL, 0);
			if (size_local_extra<0) {
				free(extrafield);
				free(commentary);
				break;
			}

			local_extra = malloc(size_local_extra);
			if ((local_extra==NULL)&&(size_local_extra!=0)) {
				free(extrafield);
				free(commentary);
				break;}
			if (unzGetLocalExtrafield(szip, local_extra, size_local_extra)<0) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				break;}

			// this malloc may fail if file very large
			buf = malloc(unzfi.uncompressed_size);
			//buf2=(char*)malloc(unzfi.uncompressed_size);
			if ((buf==NULL)&&(unzfi.uncompressed_size!=0)) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				break;}

			// read file
			sz = unzReadCurrentFile(szip, buf, unzfi.uncompressed_size);
			//modify_xml(buf, unzfi.uncompressed_size);
			if ((unsigned int)sz != unzfi.uncompressed_size) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				free(buf);
				break;}

			// open destination file
			//dest , src ,size
			memcpy(&zfi.tmz_date, &unzfi.tmu_date, sizeof(tm_unz));
			zfi.dosDate = unzfi.dosDate;
			zfi.internal_fa = unzfi.internal_fa;
			zfi.external_fa = unzfi.external_fa;

			if (zipOpenNewFileInZip2 (dzip, dos_fn, &zfi, local_extra, size_local_extra, extrafield, unzfi.size_file_extra, commentary, method, level, 0)!=UNZ_OK) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				free(buf);
				break;}

			// write file
			if (zipWriteInFileInZip(dzip, buf, modify_xml(buf, unzfi.uncompressed_size) )!=UNZ_OK) {   //17 remove cheat end
				free(extrafield);
				free(commentary);
				free(local_extra);
				free(buf);
				break;}


			//if (zipCloseFileInZipRaw(dzip, unzfi.uncompressed_size, unzfi.crc)!=UNZ_OK) {
			if (zipCloseFileInZip(dzip)!=UNZ_OK) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				free(buf);
				break;}

			if (unzCloseCurrentFile(szip)==UNZ_CRCERROR) {
				free(extrafield);
				free(commentary);
				free(local_extra); 
				free(buf);
				break;}

			free(commentary);
			free(buf);
			free(extrafield);
			free(local_extra);

			n_files ++;

			some_was_del = true;
		}
		else
		{
			extrafield = (char*)malloc(unzfi.size_file_extra);
			if ((extrafield==NULL)&&(unzfi.size_file_extra!=0))
				break;
			commentary = (char*)malloc(unzfi.size_file_comment);

			if ((commentary==NULL)&&(unzfi.size_file_comment!=0)) {
				free(extrafield);
				break;}

			if ( unzGetCurrentFileInfo(szip, &unzfi, dos_fn, MAX_PATH, extrafield, unzfi.size_file_extra, commentary, unzfi.size_file_comment) != UNZ_OK) {
				free(extrafield);
				free(commentary);
				break;}

			// open file for RAW reading
			if (unzOpenCurrentFile2(szip, &method, &level, 1)!=UNZ_OK) {
				free(extrafield);
				free(commentary);
				break;
			}

			size_local_extra = unzGetLocalExtrafield(szip, NULL, 0);
			if (size_local_extra<0) {
				free(extrafield);
				free(commentary);
				break;
			}

			local_extra = malloc(size_local_extra);
			if ((local_extra==NULL)&&(size_local_extra!=0)) {
				free(extrafield);
				free(commentary);
				break;}
			if (unzGetLocalExtrafield(szip, local_extra, size_local_extra)<0) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				break;}

			// this malloc may fail if file very large
			buf = malloc(unzfi.compressed_size);
			if ((buf==NULL)&&(unzfi.compressed_size!=0)) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				break;}

			// read file
			sz = unzReadCurrentFile(szip, buf, unzfi.compressed_size);
			if ((unsigned int)sz != unzfi.compressed_size) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				free(buf);
				break;}

			// open destination file
			//dest , src ,size
			memcpy(&zfi.tmz_date, &unzfi.tmu_date, sizeof(tm_unz));
			zfi.dosDate = unzfi.dosDate;
			zfi.internal_fa = unzfi.internal_fa;
			zfi.external_fa = unzfi.external_fa;

			if (zipOpenNewFileInZip2 (dzip, dos_fn, &zfi, local_extra, size_local_extra, extrafield, unzfi.size_file_extra, commentary, method, level, 1)!=UNZ_OK) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				free(buf);
				break;}

			// write file
			if (zipWriteInFileInZip(dzip, buf, unzfi.compressed_size)!=UNZ_OK) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				free(buf);
				break;}

			if (zipCloseFileInZipRaw(dzip, unzfi.uncompressed_size, unzfi.crc)!=UNZ_OK) {
				free(extrafield);
				free(commentary);
				free(local_extra);
				free(buf);
				break;}

			if (unzCloseCurrentFile(szip)==UNZ_CRCERROR) {
				free(extrafield);
				free(commentary);
				free(local_extra); 
				free(buf);
				break;}

			free(commentary);
			free(buf);
			free(extrafield);
			free(local_extra);

			n_files ++;
		}

		rv = unzGoToNextFile(szip);
	}

	zipClose(dzip, glob_comment);
	unzClose(szip);

	free(glob_comment);


	// if fail
	if ( (!some_was_del) || (rv!=UNZ_END_OF_LIST_OF_FILE) )
	{
		remove(tmp_name);		
		free(tmp_name);
		return 1;
	}

	remove(zip_name);
	if (rename(tmp_name, zip_name) != 0)
	{
		free(tmp_name);
		return -1;
	}

	// if all files were deleted
	if (n_files==0)
		remove(zip_name);

	free(tmp_name);
	return 0;
}