package com.wcinformatics.gmail;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.Properties;

import javax.mail.BodyPart;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.Part;
import javax.mail.Session;
import javax.mail.Store;

import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Utility for downloading SMS/MMS messages from Gmail that have been backed up
 * using SMS Backup+ (android).
 * 
 * NOTE: your gmail settings must allow this kind of access (more details
 * needed)
 */
public class SMSDownload {

	/**
	 * The main method.
	 * 
	 * @param args
	 *            the arguments - no arguments needed, configure variables
	 * @throws IOException
	 *             Signals that an I/O exception has occurred.
	 */
	public static void main(String args[]) throws IOException {
		Properties props = System.getProperties();
		props.setProperty("mail.store.protocol", "imaps");
		// This should match the name of the person in the subject of the therad
		String person = "John Smith";
		// This is the output file
		String file = "c:/Users/Jane Doe/Desktop/output.xls";
		// The gmail username and password
		String username = "jane.doe@gmail.com";
		String password = "p#ssw0rd";
		try {

			// Set up imap session and properties
			Session session = Session.getDefaultInstance(props, null);
			Store store = session.getStore("imaps");
			// Here is the username and password
			store.connect("imap.gmail.com", username, password);
			System.out.println("connection = store");

			// Set up workbook
			Workbook workbook = new HSSFWorkbook();
			CreationHelper createHelper = workbook.getCreationHelper();
			Sheet sheet = workbook.createSheet(person);
			Font font = workbook.createFont();
			font.setFontName("Calibri");
			font.setFontHeightInPoints((short) 11);
			CellStyle style = workbook.createCellStyle();
			style.setFont(font);
			CellStyle dateStyle = workbook.createCellStyle();
			dateStyle.setDataFormat(createHelper.createDataFormat().getFormat(
					"MM/dd/yyyy HH:mm:ss"));

			// Create the drawing patriarch. This is the top level container for
			// all shapes.
			Drawing drawing = sheet.createDrawingPatriarch();

			// Set up headers
			int rownum = 0;
			Row row = sheet.createRow(rownum);
			int cellnum = 0;
			Cell cell = row.createCell(cellnum++);
			cell.setCellStyle(style);
			cell.setCellValue(createHelper.createRichTextString("Date"));
			cell = row.createCell(cellnum++);
			cell.setCellStyle(style);
			cell.setCellValue(createHelper.createRichTextString("From"));
			cell = row.createCell(cellnum++);
			cell.setCellStyle(style);
			cell.setCellValue(createHelper.createRichTextString("Text"));
			cell = row.createCell(cellnum++);
			cell.setCellStyle(style);
			cell.setCellValue(createHelper.createRichTextString("MMS"));

			// Get messages
			Folder smsFolder = store.getFolder("SMS");
			smsFolder.open(Folder.READ_ONLY);
			Message messages[] = smsFolder.getMessages();
			for (Message message : messages) {
				if (message.getSubject().equals("SMS with " + person)) {
					row = sheet.createRow(++rownum);
					cellnum = 0;
					cell = row.createCell(cellnum++);
					cell.setCellStyle(dateStyle);
					Date date = message.getReceivedDate();
					cell.setCellValue(createHelper.createRichTextString(date
							.toString()));
					cell = row.createCell(cellnum++);
					cell.setCellStyle(style);
					cell.setCellValue(createHelper.createRichTextString(message
							.getFrom()[0].toString()));
					cell = row.createCell(cellnum++);
					cell.setCellStyle(style);
					cell.setCellValue(createHelper.createRichTextString(message
							.getContent().toString()));

					if (message.getContent() instanceof Multipart) {
						Multipart multipart = (Multipart) message.getContent();
						for (int i = 0; i < multipart.getCount(); i++) {
							BodyPart bodyPart = multipart.getBodyPart(i);
							if (!Part.ATTACHMENT.equalsIgnoreCase(bodyPart
									.getDisposition())) {
								continue; // dealing with attachments only
							}
							if (!bodyPart.isMimeType("image/jpeg")) {
								System.out.println("Skip "
										+ bodyPart.getContentType());
								continue;
							}
							System.out.println("Handle "
									+ bodyPart.getContentType());
							InputStream is = bodyPart.getInputStream();
							// add picture data to this workbook.
							byte[] bytes = IOUtils.toByteArray(is);
							int pictureIdx = workbook.addPicture(bytes,
									Workbook.PICTURE_TYPE_JPEG);
							is.close();

							// add a picture shape
							ClientAnchor anchor = createHelper
									.createClientAnchor();
							// set top-left corner of the picture,
							// subsequent call of Picture#resize() will operate
							// relative to it
							anchor.setCol1(cellnum);
							anchor.setRow1(rownum);
							// 0 = Move and size with Cells, 2 = Move but don't
							// size with cells, 3 = Don't move or size with
							// cells.
							anchor.setAnchorType(2);

							Picture pict = drawing.createPicture(anchor,
									pictureIdx);
							pict.resize();
						}
					}
					System.out.println(date + " " + message.getFrom()[0] + ", "
							+ message.getContent());
				}
			}

			System.out.println("output file = " + file);
			FileOutputStream out = new FileOutputStream(new File(file));
			workbook.write(out);
			System.out.println("  DONE");
		} catch (Exception e) {
			e.printStackTrace();
			System.exit(2);
		}
	}
}
