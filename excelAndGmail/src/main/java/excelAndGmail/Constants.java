package excelAndGmail;

import java.util.Collections;
import java.util.List;

import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.gmail.GmailScopes;

public class Constants {

	//プログラムで使用する定数を定義
	static String PATH = "/Users/ExcelFile/";

	static final String APPLICATON_NAME = "GmailApi";

	static final String USER = "me";

	//変更不可の単一要素を持つListの作成
	static final List<String> SCOPES = Collections.singletonList(GmailScopes.GMAIL_READONLY);

	//JacksonでJSONファイルをJavaで使用可能にする
	static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

	static final String CLIENT_SECRET_DIR = "client_secret_946481271572-7055q1b2og4l9mpjmaukf4ih5fkd7k0l.apps.googleusercontent.com.json";

	static final String CREDENTIALS_FOLDER = "credentials";

	static final String MEMBER = "Your addressi";
}
