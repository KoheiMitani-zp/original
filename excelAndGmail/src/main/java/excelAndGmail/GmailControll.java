package excelAndGmail;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.lang3.time.DateUtils;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.gmail.Gmail;
import com.google.api.services.gmail.model.ListMessagesResponse;
import com.google.api.services.gmail.model.Message;
import com.google.api.services.gmail.model.MessagePartHeader;

//Gmailを操作するクラス
public class GmailControll {

	public Map<String, String> gmailConnect(List<String> mailList) throws GeneralSecurityException, IOException{

		Date day = DateUtils.truncate(new Date(), Calendar.DAY_OF_MONTH);

		Calendar cal = Calendar.getInstance();

		//取得するメールを日付で選定するために必要
		SimpleDateFormat  dateFormat = new SimpleDateFormat("EEE, d MMM yyyy", Locale.ENGLISH);
		//
		SimpleDateFormat dateFormatForQuery = new SimpleDateFormat("yyyy/MM/dd");

		cal.setTime(day);
		//対象とするメールの日付を取得
		cal.add(Calendar.DATE, -1);

		day = cal.getTime();

		//メールを取得する際にAPIで使用するクエリとメールheadersに使用するための変数を用意（どちらか一つでもおそらく大丈夫）
		String dateF = dateFormat.format(day);
		String dateQ = dateFormatForQuery.format(day);

		//Http通信
		NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();

		//秘密鍵の取得
		FileInputStream in = new FileInputStream(new File(Constants.CLIENT_SECRET_DIR));
		InputStreamReader reader = new InputStreamReader(in);
		GoogleClientSecrets clientsecrets = GoogleClientSecrets.load(Constants.JSON_FACTORY, reader);

		//GmailAPI認証
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, Constants.JSON_FACTORY, clientsecrets, Constants.SCOPES)
				.setDataStoreFactory(new FileDataStoreFactory(new java.io.File(Constants.CREDENTIALS_FOLDER)))
				.setAccessType("offline").build();

		//認証の実施
		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");

		//メールをすべて取得
		Gmail service = new Gmail.Builder(HTTP_TRANSPORT, Constants.JSON_FACTORY, credential).setApplicationName(Constants.APPLICATON_NAME).build();

		//取得したメールからメールを絞るための条件クエリ（検索演算子を使用）
		String query = "after:" + dateQ + "OR subject:GmailApiテストメール";

		//条件クエリに基づいて、メッセ-ジを取得を実行（メールのIdのみだけ取得）
		ListMessagesResponse messagesResponse = service.users().messages().list(Constants.USER).setQ(query).execute();
		List<Message> messageList = messagesResponse.getMessages();


		//対象のメールを見つけるために取得したメールのIdとともに、ほかの全データを取得
		Message targetMessage = null;
		for(Message message : messageList) {
			//メールの全データを取得
			Message fullMessage = service.users().messages().get(Constants.USER, message.getId()).execute();
			//ヘッダーの取得
			for(MessagePartHeader header : fullMessage.getPayload().getHeaders()) {
				//ヘッダーの中で件名がGmailApiを含むメールの全データを取得（検索条件で一度絞っていれば、書かなくてよい）
				if(header.getName().equals("Subject") && header.getValue().contains("GmailApiテストメール")) {
					targetMessage = fullMessage;
				}
			}
		}

		//送信元メールアドレスと、送信先メールアドレスをセットで格納するためのマップ
		String targetMailAddress = null;
		String toTargetMessageAddress = null;
		Map<String, String> targetAddressSet = new HashMap<String, String>();

		String bodyBase64 = null;

		//対象メールの本文を取得（Base64）、読めるようにUTF-8にまでデコードする必要有
		//メールを複数取得する際は、必要なデータが入っている場所が違うことがあるので条件分岐が必要な時有
		bodyBase64 = targetMessage.getPayload().getParts().get(0).getBody().get("data").toString();
		byte[] bodyBytes = Base64.decodeBase64(bodyBase64);
		String body = new String(bodyBytes, "UTF-8");

		System.out.println(body);

		//本文に『テストメールだぜ！！』が含まれていれば、送信元メールアドレスと送信先メールアドレスを取得
		if(body.contains("テストメールだぜ")) {
			for(int i = 0; i < targetMessage.getPayload().getHeaders().size(); i++) {
				if(targetMessage.getPayload().getHeaders().get(i).getName().equals("From")) {
					targetMailAddress =	targetMessage.getPayload().getHeaders().get(i).getValue();
				}
			}for(int i = 0; i < targetMessage.getPayload().getHeaders().size(); i++) {
				if(targetMessage.getPayload().getHeaders().get(i).getName().equals("To")) {
					toTargetMessageAddress = targetMessage.getPayload().getHeaders().get(i).getValue();
				}
			}
		}
		targetAddressSet.put(targetMailAddress, toTargetMessageAddress);
		for(Map.Entry<String, String> entry : targetAddressSet.entrySet()) {
			System.out.println(entry.getKey() + " : " + entry.getValue());
		}
		return targetAddressSet;
	}
}
