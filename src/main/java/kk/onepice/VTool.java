package kk.onepice;

import com.sun.javafx.tk.FileChooserType;
import it.sauronsoftware.jave.Encoder;
import it.sauronsoftware.jave.EncoderException;
import it.sauronsoftware.jave.MultimediaInfo;
import it.sauronsoftware.jave.VideoSize;
import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Pane;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class VTool extends Application {
	TextArea textArea;
	File targetFolder;
	Map<String, VideoSize> fileInfo;

	File sendEmailFile;

	public static void main(String[] args) {
		launch(args);
	}


	/**
	 * 初始化处理视频的UI控件
	 */
	private void initHandVideoUI(Stage stage,GridPane inputGridPane){
		DirectoryChooser videoDirectoryChooser = new DirectoryChooser();
		videoDirectoryChooser.setTitle("选择文件夹");

		Button openFolderButton = new Button("选择视频目录");
		Label folderLabel = new Label("未指定视频目录");
		folderLabel.setPrefWidth(200);
		openFolderButton.setOnAction(
				(final ActionEvent e) -> {
					targetFolder = videoDirectoryChooser.showDialog(stage);
					if (targetFolder != null) {
						folderLabel.setText(targetFolder.getPath());
					}
				});

		final Button startButton = new Button("开始处理视频");
		startButton.setOnAction(
				(final ActionEvent e) -> {
					if (targetFolder == null) {
						Alert alert = new Alert(Alert.AlertType.WARNING);
						alert.setTitle("警告");
						alert.setHeaderText("请指定需要处理的视频目录！");
						alert.showAndWait();
						return;
					}
					startHandle();
				});

		GridPane.setConstraints(openFolderButton, 0, 0);
		GridPane.setConstraints(folderLabel, 1, 0);
		GridPane.setConstraints(startButton, 2, 0);
		inputGridPane.getChildren().addAll(openFolderButton, folderLabel, startButton);
	}

	/**
	 * 初始化处理邮件的UI控件
	 */
	private void initHandEmailUI(Stage stage,GridPane inputGridPane){
		FileChooser emailFileChooser = new FileChooser();
		emailFileChooser.setTitle("选择文件");

		Button openFileChooserButton = new Button("选择表格文件");
		Label fileChooserLabel = new Label("未指定文件");

		openFileChooserButton.setOnAction(
				(final ActionEvent e) -> {
					sendEmailFile = emailFileChooser.showOpenDialog(stage);
					if (sendEmailFile != null) {
						fileChooserLabel.setText(sendEmailFile.getPath());
					}
				});

		final Button sendMailButton = new Button("开始发送邮件");
		sendMailButton.setOnAction(
				(final ActionEvent e) -> {
					if (sendEmailFile == null) {
						Alert alert = new Alert(Alert.AlertType.WARNING);
						alert.setTitle("警告");
						alert.setHeaderText("请指定发送邮件表格！");
						alert.showAndWait();
						return;
					}
					try {
						startSendMail();
					}catch (Exception ex){
						printlnLog(ex.toString());
					}

				});

		GridPane.setConstraints(openFileChooserButton, 0, 1);
		GridPane.setConstraints(fileChooserLabel, 1, 1);
		GridPane.setConstraints(sendMailButton, 2, 1);
		inputGridPane.getChildren().addAll(openFileChooserButton, fileChooserLabel, sendMailButton);
	}

	private void startSendMail() throws Exception{
		InputStream inputStream = new FileInputStream(sendEmailFile);
		XSSFWorkbook book = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = book.getSheetAt(0);
		int maxRowNum = sheet.getLastRowNum() + 1;
		for (int i = 0; i < maxRowNum; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				String eAddr= row.getCell(0).getStringCellValue();
				String eContent = row.getCell(1).getStringCellValue();
				if(StringUtils.isBlank(eAddr)||StringUtils.isBlank(eContent)){
					continue;
				}
				printlnLog(eAddr+":"+eContent);
			}
		}

	}

	@Override
	public void start(final Stage stage) {
		final GridPane inputGridPane = new GridPane();

		// 处理视频功能
		initHandVideoUI(stage,inputGridPane);
		// 处理邮件功能
		initHandEmailUI(stage,inputGridPane);

		// 初始化日志打印区
		textArea = new TextArea();
		textArea.setPrefSize(400, 300);
		textArea.setEditable(false);
		GridPane.setConstraints(textArea, 0, 2, 3, 1);

		inputGridPane.setHgap(6);
		inputGridPane.setVgap(6);
		inputGridPane.getChildren().add(textArea);

		final Pane root = new VBox(12);
		root.getChildren().add(inputGridPane);
		root.setPadding(new Insets(12, 12, 12, 12));
		Scene scene = new Scene(root, 500, 400);
		stage.setScene(scene);
		stage.setTitle("VTool");
		stage.setResizable(false);
		stage.show();
	}

	private void startHandle() {
		printlnLog("-------------------------------------------");
		printlnLog("开始处理...");
		fileInfo = new HashMap<>();
		if (targetFolder.exists()) {
			handleFolder(targetFolder);
		}
		saveFileInfo();
	}

	private void saveFileInfo() {
		//创建一个表格对象
		Workbook workbook = new XSSFWorkbook();
		String[] s = new String[]{"视频名", "宽高"};
		Sheet sheet = workbook.createSheet();
		int rowNum = 0;
		//设置标题
		Row head = sheet.createRow(rowNum++);
		for (int i = 0; i < s.length; i++) {
			Cell cell = head.createCell(i);
			cell.setCellValue(s[i]);
		}
		//设置内容
		for (String key : fileInfo.keySet()) {
			Row row = sheet.createRow(rowNum++);
			Cell cell0 = row.createCell(0);
			cell0.setCellValue(key);
			Cell cell1 = row.createCell(1);
			cell1.setCellValue(fileInfo.get(key).getWidth() + "*" + fileInfo.get(key).getHeight());
		}
		//创建文件
		File file = new File("/" + System.currentTimeMillis() + ".xlsx");
		//输出到指定位置
		FileOutputStream outputStream = null;
		try {
			outputStream = new FileOutputStream(file);
			workbook.write(outputStream);
		} catch (IOException e) {
		} finally {
			try {
				outputStream.close();
			} catch (IOException e) {
			}
		}
		printlnLog("保存文件信息成功:" + file.getAbsolutePath());
	}

	/**
	 * 递归遍历文件夹
	 */
	private void handleFolder(File folder) {
		File[] files = folder.listFiles();
		for (File file : files) {
			if (file.isDirectory()) {
				handleFolder(file);
			} else {
				handleFile(file);
			}
		}
	}

	/**
	 * 处理文件
	 */
	private void handleFile(File file) {
		String fileName = file.getName();
		String fileType = fileName.substring(fileName.lastIndexOf(".") + 1);
		String path = file.getPath();
		if (fileType.equalsIgnoreCase("mov") || fileType.equalsIgnoreCase("mp4")) {
			Encoder encoder = new Encoder();
			try {
				MultimediaInfo multimediaInfo = encoder.getInfo(file);
				VideoSize videoSize = multimediaInfo.getVideo().getSize();
				fileInfo.put(fileName, videoSize);
				printlnLog("Info：获取视频帧尺寸 " + fileName + "[" + videoSize.getWidth() + "*" + videoSize.getHeight() + "]");
			} catch (EncoderException e) {
				e.printStackTrace();
			}
			if (fileType.equalsIgnoreCase("mov")) {
				file.setWritable(true);
				boolean b = file.renameTo(new File(path.substring(0, path.lastIndexOf(".")) + ".mp4"));
				if (b) {
					printlnLog("Info：修改文件名成功 " + fileName);
				} else {
					printlnLog("Error：修改文件名失败，请关闭文件及其所在目录后重新尝试 " + fileName);
				}
			}
		} else {
			printlnLog("Warning：无法识别的文件 " + fileName);
		}
	}

	private void printlnLog(String str) {
		textArea.appendText(str + "\n");
		textArea.setScrollTop(textArea.getScrollTop());
	}

}
