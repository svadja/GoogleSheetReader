package ua.svadja.googleSheetReader;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.List;

import javax.annotation.PostConstruct;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

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
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.ValueRange;

/**
 * Created by vsasa
 */
@Component
public class GoogleSpreadsheet {

    private static final String APPLICATION_NAME = "Camunda BPM";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS);
    private static final String CALLBACK_URL = "http://localhost:8080/Callback";

    private HttpTransport httpTransport;

    @Value("${security.google.clientSecret}")
    private String clientSecret;
    @Value("${security.google.googleDataStore}")
    private String googleDataStore;

    private Credential credential;
    private Sheets sheetsService;

    public GoogleSpreadsheet() {
    }

    @PostConstruct
    private void setup() {
        try {
            httpTransport = GoogleNetHttpTransport.newTrustedTransport();
            this.credential = authorize();
            this.sheetsService = getSheetsService();
        } catch (Exception e) {
            e.printStackTrace();
            System.exit(1);
        }
    }

    private Credential authorize() throws IOException, GeneralSecurityException {
        InputStream in = getClass().getResourceAsStream(clientSecret);
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));
        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow =
                new GoogleAuthorizationCodeFlow.Builder(
                        httpTransport, JSON_FACTORY, clientSecrets, SCOPES)
                        .setDataStoreFactory(new FileDataStoreFactory(new File(googleDataStore)))
                        .setApprovalPrompt("force")
                        .setAccessType("offline")
                        .build();
/* //Next code uses manual request code without server receiver

        String authorizeUrl = flow.newAuthorizationUrl().setRedirectUri(CALLBACK_URL).setApprovalPrompt("force")
                .setAccessType("offline").build();
        System.out.printf("Paste this url in your browser:%n%s%n", authorizeUrl);
        System.out.println("Type the code you received here: ");
        String authorizationCode = new BufferedReader(new InputStreamReader(System.in)).readLine();
        GoogleAuthorizationCodeTokenRequest tokenRequest = flow.newTokenRequest(authorizationCode);
        tokenRequest.setRedirectUri(CALLBACK_URL);
        GoogleTokenResponse tokenResponse = tokenRequest.execute();
        GoogleCredential credential = new GoogleCredential.Builder()
                .setTransport(new NetHttpTransport())
                .setJsonFactory(new JacksonFactory())
                .setClientSecrets(clientSecrets)
                .build().setFromTokenResponse(tokenResponse);
*/
        LocalServerReceiver localServerReceiver = new LocalServerReceiver.Builder().setPort(8080).build();
        Credential credential =  new AuthorizationCodeInstalledApp(flow, localServerReceiver).authorize("user");

        return credential;
    }

    private Sheets getSheetsService() throws IOException {
        return new Sheets.Builder(httpTransport, JSON_FACTORY, this.credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    /**
     * @param spreadsheetId - from https://docs.google.com/spreadsheets/d/spreadsheetId/
     * @param range         -
     *                      Sheet1!A1:B2 refers to the first two cells in the top two rows of Sheet1.
     *                      Sheet1!A:A refers to all the cells in the first column of Sheet1.
     *                      Sheet1!1:2 refers to the all the cells in the first two rows of Sheet1.
     *                      Sheet1!A5:A refers to all the cells of the first column of Sheet 1, from row 5 onward.
     *                      A1:B2 refers to the first two cells in the top two rows of the first visible sheet.
     *                      Sheet1 refers to all the cells in Sheet1.
     * @return
     * @throws IOException
     */
    public List<List<Object>> getData(String spreadsheetId, String range) throws IOException{
        return sheetsService.spreadsheets().values().get(spreadsheetId, range).execute().getValues();
    }

    /**
     * Get data from page
     *
     * @param spreadsheetId - from https://docs.google.com/spreadsheets/d/spreadsheetId/
     * @param sheetNumber   - start from 0
     * @param cellsRange    - like A2:B
     * @return
     * @throws IOException
     */
    public List<List<Object>> getData(String spreadsheetId, int sheetNumber, String cellsRange) throws IOException {
        List<Sheet> sheets = sheetsService.spreadsheets().get(spreadsheetId).execute().getSheets();
        String cellsRangeStringPart = "";
        if (!cellsRange.isEmpty()) {
            cellsRangeStringPart = ":" + cellsRange;
        }
        return sheetsService.spreadsheets().values().get(spreadsheetId, sheets.get(sheetNumber).getProperties().getTitle() + cellsRangeStringPart).execute().getValues();
    }

    public void appendData(String spreadsheetId, String range, List<List<Object>> values) throws IOException{
        ValueRange valueRange = new ValueRange();
        valueRange.setValues(values);
        sheetsService.spreadsheets().values().append(spreadsheetId, range, valueRange).setValueInputOption("RAW").execute();
    }

    public void updateData(String spreadsheetId, String range, List<List<Object>> values) throws IOException{
        ValueRange valueRange = new ValueRange();
        valueRange.setValues(values);
        sheetsService.spreadsheets().values().update(spreadsheetId, range, valueRange).execute();
    }
}
