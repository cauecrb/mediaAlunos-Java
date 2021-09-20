import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.apiclient.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.time.temporal.ValueRange;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;


// class Media Alunos
public class MediaAlunos {
    private static Sheets sheetsService;
    private static final String APPLICATION_NAME = "Google Sheets API Java MediaAlunos";
    private static String SPREADSHEET_ID "1GBXK7bcZwdJTtGr4AGn9fFloLcqq3LsrgTadEgg6UQo";


    //authorize for google sheets
    private  static Credential authorize() throws IOException, GeneralSecurityException {
        InputStream in = MediaAlunos.class.getResourceAsStream("/credentials.json");
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(
                JacksonFactory.getDefaultInstance(), new InputStreamReader(in)
        );

        List<String> scopes = Arrays.asList(SheetsScopes.SPREADSHEETS);

        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                GoogleNetHttpTransport.newTrustedTransport(), JacksonFactory.getDefaultInstance(),
                clientSecrets, scopes)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File("tokens")))
                .setAccesType("offline").build();

        Credential credential = new AuthorizationCodeInstallApp(
                flow, new LocalServerReceiver())
                .authorize("user");
        return credential;
    }

    //get google Sheets aplication

    public static Sheets getSheetsService() throws IOException, GeneralSecurityException{
        Credential credential = authorize();
        return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(),
                JacksonFactory.getDefaultInstance(), credential)
                .setApplicationName(APPLICATION_NAME).build();

    }


    // verify if aproved or not and set in table
    public static void main(String[] args) throws IOException, GeneralSecurityException {
        sheetsService = getSheetsService();
        for(int i = 4; i < 28; i++ ){
            String range = "Congres!A"+i+":F"+i;

            string row1 = "G"+i;
            string row2 = "H"+i;

            ValueRange response = sheetsService.spreadsheets().values()
                    .get(SPREADSHEET_ID, range).execute();

            List<List<Object>> values = response.getValues();

            if(values == null || values.isEmpty()){
                System.out.println("fim do documento");
            } else{

                string body;

                if(row.get(2) > 15){

                    body = "Reprovado por Falta";
                    UpdateValuesResponse result = sheetsService.spreadsheets().values()
                            .update(SPREADSHEET_ID,row1,body)
                            .setValueInputOption("RAW").execute();
                    body = 0;
                    UpdateValuesResponse result = sheetsService.spreadsheets().values()
                            .update(SPREADSHEET_ID,row2,body)
                            .setValueInputOption("RAW").execute();
                }

                int media = row.get(3) + row.get(4) + row.get(5);
                media = media/3;

                if(media >= 7 ){

                    body = "Aprovado";
                    UpdateValuesResponse result = sheetsService.spreadsheets().values()
                            .update(SPREADSHEET_ID,row1,body)
                            .setValueInputOption("RAW").execute();
                    body = 0;
                    UpdateValuesResponse result = sheetsService.spreadsheets().values()
                            .update(SPREADSHEET_ID,row2,body)
                            .setValueInputOption("RAW").execute();
                }

                if(media < 5 ){

                    body = "Reprovado por Nota";
                    UpdateValuesResponse result = sheetsService.spreadsheets().values()
                            .update(SPREADSHEET_ID,row1,body)
                            .setValueInputOption("RAW").execute();
                    body = "0";
                    UpdateValuesResponse result = sheetsService.spreadsheets().values()
                            .update(SPREADSHEET_ID,row2,body)
                            .setValueInputOption("RAW").execute();
                }

                else{

                    body = "Exame Final";
                    UpdateValuesResponse result = sheetsService.spreadsheets().values()
                            .update(SPREADSHEET_ID,row1,body)
                            .setValueInputOption("RAW").execute();
                    body = 10 - media;
                    UpdateValuesResponse result = sheetsService.spreadsheets().values()
                            .update(SPREADSHEET_ID,row2,body)
                            .setValueInputOption("RAW").execute();
                }
            }
        }
    }

}
