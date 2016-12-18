package com.intumit.ryanchen.googleSheets;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;

import javax.swing.undo.CannotUndoException;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetResponse;
import com.google.api.services.sheets.v4.model.CellData;
import com.google.api.services.sheets.v4.model.CellFormat;
import com.google.api.services.sheets.v4.model.ExtendedValue;
import com.google.api.services.sheets.v4.model.GridCoordinate;
import com.google.api.services.sheets.v4.model.GridRange;
import com.google.api.services.sheets.v4.model.NumberFormat;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.RowData;
import com.google.api.services.sheets.v4.model.UpdateCellsRequest;

public class GoogleSheets {
	/** Application name. */
	private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";

	/** Directory to store user credentials for this application. */
	private static final java.io.File DATA_STORE_DIR = new java.io.File(System.getProperty("user.home"),
			".credentials/sheets.googleapis.com-java-quickstart");

	/** Global instance of the {@link FileDataStoreFactory}. */
	private static FileDataStoreFactory DATA_STORE_FACTORY;

	/** Global instance of the JSON factory. */
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

	/** Global instance of the HTTP transport. */
	private static HttpTransport HTTP_TRANSPORT;

	/**
	 * Global instance of the scopes required by this quickstart.
	 *
	 * If modifying these scopes, delete your previously saved credentials at
	 * ~/.credentials/sheets.googleapis.com-java-quickstart
	 */
	private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS);

	static {
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
		} catch (Throwable t) {
			t.printStackTrace();
			System.exit(1);
		}
	}

	/**
	 * Creates an authorized Credential object.
	 * 
	 * @return an authorized Credential object.
	 * @throws IOException
	 */
	public static Credential authorize() throws IOException {
		// Load client secrets.
		// InputStream in = new
		// FileInputStream("C:\\Users\\UtahC\\.credentials\\sheets.googleapis.com-java-quickstart\\client_secret.json");
		InputStream in = new FileInputStream(".\\client_secret.json");
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		// Build flow and trigger user authorization request.
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY,
				clientSecrets, SCOPES).setDataStoreFactory(DATA_STORE_FACTORY).setAccessType("offline").build();
		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
		System.out.println("Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
		return credential;
	}

	/**
	 * Build and return an authorized Sheets API client service.
	 * 
	 * @return an authorized Sheets API client service
	 * @throws IOException
	 */
	public static Sheets getSheetsService() throws IOException {
		Credential credential = authorize();
		return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential).setApplicationName(APPLICATION_NAME)
				.build();
	}

	public static void main(String[] args) throws IOException {
		
        Sheets service = getSheetsService();
        String spreadsheetId = "12KNDn5Jiwbl3YbxTgXURofS6wwhTXGQn2JilxZ9DnDM";
        
    	List<Request> requests = new ArrayList<>();
//    	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("test2")));
//    	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("test3")));
//    	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("test4")));
//    	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("test5")));
//    	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("test6")));
    	requests.add(updateCellRequest(0, 5, 0, new Date()));
//    	requests.add(new Request()
//                .setUpdateCells(new UpdateCellsRequest()
//                        .setRange(new GridRange()
//                        		.setSheetId(0)
//                        		.setStartRowIndex(20)
//                        		.setStartColumnIndex(0))
//                        .setRows(Arrays.asList(
//                                new RowData().setValues(values)))
//                        .setFields("userEnteredValue,userEnteredFormat.backgroundColor")));
    	
    	BatchUpdateSpreadsheetRequest body =
    	        new BatchUpdateSpreadsheetRequest().setRequests(requests);
    	BatchUpdateSpreadsheetResponse response =
    	        service.spreadsheets().batchUpdate(spreadsheetId, body).execute();
    	System.out.println(response);
    }

	static String readFile(String path, Charset encoding) throws IOException {
		byte[] encoded = Files.readAllBytes(Paths.get(path));
		return new String(encoded, encoding);
	}
	
	static Request updateCellRequest(int sheetId, int rowIndex, int colIndex, Object value) {
		Request request = new Request()
			      .setUpdateCells(new UpdateCellsRequest()
			              .setStart(new GridCoordinate()
			                      .setSheetId(0)
			                      .setRowIndex(rowIndex)
			                      .setColumnIndex(colIndex))
			              .setFields("userEnteredValue,userEnteredFormat.backgroundColor"));
		
		List<CellData> values = new ArrayList<>();
		if (value.getClass() == Date.class) {
			values.add(new CellData()
					.setUserEnteredValue(new ExtendedValue()
							.setNumberValue(getSerialNumber(new Date())))
					.setUserEnteredFormat(new CellFormat()
							.setNumberFormat(new NumberFormat().setType("DATE"))));
			
			UpdateCellsRequest updateCellsRequest = request.getUpdateCells();
			updateCellsRequest.setFields(updateCellsRequest.getFields() + ",userEnteredFormat.numberFormat");
		}
		else {
			values.add(new CellData()
					.setUserEnteredValue(new ExtendedValue()
							.setStringValue(value.toString())));
		}
		
		request.getUpdateCells().setRows(Arrays.asList(
                new RowData().setValues(values)));
        
		
		return request;
	}
	
	static double getSerialNumber(Date date) {
		Calendar c1 = Calendar.getInstance();
	    c1.set(1899, 12, 30, 0, 0, 0);
	    Calendar c2 = Calendar.getInstance();
	    c2.setTime(date);
	    System.out.println("c1.getTimeInMillis() = " + c1.getTimeInMillis());
	    System.out.println("c2.getTimeInMillis() = " + c2.getTimeInMillis());
	    long diff = (c2.getTimeInMillis() - c1.getTimeInMillis());
	    System.out.println("TimeUnit.DAYS.convert(diff, TimeUnit.MICROSECONDS) = " + TimeUnit.DAYS.convert(diff, TimeUnit.MICROSECONDS));
	    System.out.println("TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS) = " + TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS));
	    
	    return diff / 1000d / 60 / 60 / 24;
	}

}