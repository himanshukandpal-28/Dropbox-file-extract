package com.demo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import com.dropbox.core.DbxApiException;
import com.dropbox.core.DbxException;
import com.dropbox.core.DbxRequestConfig;
import com.dropbox.core.v2.DbxClientV2;
import com.dropbox.core.v2.files.FileMetadata;
import com.dropbox.core.v2.files.FolderMetadata;
import com.dropbox.core.v2.files.ListFolderResult;
import com.dropbox.core.v2.files.Metadata;
import com.dropbox.core.v2.sharing.CreateSharedLinkWithSettingsErrorException;
import com.dropbox.core.v2.sharing.ListSharedLinksErrorException;
import com.dropbox.core.v2.sharing.ListSharedLinksResult;
import com.dropbox.core.v2.sharing.SharedLinkMetadata;
import com.dropbox.core.v2.users.FullAccount;

/**
 * @author Himanshu Kandpal
 **/

public class DropBoxAccess {

	private static final String ACCESS_TOKEN = "sl.Bbr-L8J94KAWfreGQKe6CXDCndHEpHnws2vafJRKzxpl6COIOWKQmhfJpLn1RG4_GQ6g1BnihCyEOVyDq-wE5TC4WNyUMnFi9HWTW2mQ1DqMrX6wyaJYQP1wHbkEWAhk9kCZ7g_MTqGc";

	// Get extensions for all the file names

	private static String getExtension(Metadata metadata) {
		String ext = null;
		if (metadata instanceof FileMetadata) {
			ext = metadata.getName().substring(metadata.getName().lastIndexOf(".") + 1);
		} else if (metadata instanceof FolderMetadata) {
			ext = null;
		}

		return ext;
	}

	// Get share link of all files in the dropbox (if not avail, then create it and
	// get it)

	private static String getShareLink(DbxClientV2 client, Metadata metadata)
			throws ListSharedLinksErrorException, DbxException {
		String shareLink = null;
		try {

			// following line is to creat the share link when it does not exist!

			shareLink = client.sharing().createSharedLinkWithSettings(metadata.getPathDisplay()).getUrl();

		} catch (CreateSharedLinkWithSettingsErrorException ex) {

			// to get the share link if it already exists!

			ListSharedLinksResult result = client.sharing().listSharedLinksBuilder().withPath(metadata.getPathDisplay())
					.withDirectOnly(true).start();
			List<SharedLinkMetadata> shareLinkList = result.getLinks();
			for (SharedLinkMetadata sharedLinkMetadata : shareLinkList) {
				shareLink = sharedLinkMetadata.getUrl();
			}

		} catch (DbxException ex) {
			System.out.println(ex);
		}
		return shareLink;
	}

	/**
	 * @param metadata
	 * @return
	 */
	private static String getFileName(Metadata metadata) {
		String fileName = null;
		if (metadata instanceof FileMetadata) {
			String fullFileName = metadata.getName().substring(metadata.getName().lastIndexOf("/") + 1);
			int endIndex = fullFileName.lastIndexOf(".");
			fileName = fullFileName.substring(0, endIndex);
		}
		return fileName;
	}

	public static void main(String[] args) throws DbxApiException, DbxException, IOException {

		// Create Dropbox client
		DbxRequestConfig config = DbxRequestConfig.newBuilder("dropbox/java-tutorial").build();
		DbxClientV2 client = new DbxClientV2(config, ACCESS_TOKEN);

		FullAccount account = client.users().getCurrentAccount();
		System.out.println(account.getName().getDisplayName());

		// Get files and folder metadata from Dropbox root directory
		ListFolderResult result = client.files().listFolderBuilder("").withRecursive(true).start();
		List<JSONObject> listOfJsonObject = new ArrayList<>();
		int i = 1, j = 1;
		while (true) {
			for (Metadata metadata : result.getEntries()) {
				System.out.println("----------------------- total count ----------------------- " + i++);

				if (getFileName(metadata) != null && getExtension(metadata) != null
						&& getShareLink(client, metadata) != null) {
					JSONObject jsonObject = new JSONObject();
					jsonObject.put("path", metadata.getPathDisplay());
					jsonObject.put("file_name", getFileName(metadata));
					jsonObject.put("file_ext", getExtension(metadata));
					jsonObject.put("share_link", getShareLink(client, metadata));

					listOfJsonObject.add(jsonObject);
					System.out.println("----------------------- file count ----------------------- " + j++);
				}

			}

			if (!result.getHasMore()) {
				break;
			}
			result = client.files().listFolderContinue(result.getCursor());
		}

		// Store the files into Excel file

		String excelFile = "/Users/himanshukandpal/Desktop/Project/upwork/Dropbox-Files-Check-Demo/createdExcel1.xlsx";

		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet("Sheet1");

		int rowNum = 0;

		Row row0 = sheet.createRow(0); // add 0-th row manually!

		row0.createCell(0).setCellValue("path");
		row0.createCell(1).setCellValue("file_name");
		row0.createCell(2).setCellValue("file_ext");
		row0.createCell(3).setCellValue("share_link");

		for (JSONObject jsonObject : listOfJsonObject) {
			Row row = sheet.createRow(++rowNum);
			int colNum = 0;

			row.createCell(colNum++).setCellValue(jsonObject.getString("path"));
			row.createCell(colNum++).setCellValue(jsonObject.getString("file_name"));
			row.createCell(colNum++).setCellValue(jsonObject.getString("file_ext"));
			row.createCell(colNum++).setCellValue(jsonObject.getString("share_link"));
		}

		FileOutputStream outputStream = new FileOutputStream(excelFile);
		book.write(outputStream);

		book.close();

	}

}
