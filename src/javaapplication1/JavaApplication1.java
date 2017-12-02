/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package javaapplication1;

import com.aspose.words.Document;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigInteger;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.imageio.ImageIO;
import net.sf.json.JSONObject;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

/**
 *
 * @author zhuqiangye
 */


public class JavaApplication1 {

    private Map<String, String> key_content_map = null;
    private boolean isPreview = false;
    private XmlCursor cursor;
    private XWPFParagraph paragraph;
    private List<String> replace_values = new ArrayList<>();
    private String placehoder = "###";

    private Map getKey_value(File excel) throws FileNotFoundException, IOException {
        Map<String, String> key_content_map = new HashMap<>();
        FileInputStream xlsx_in = new FileInputStream(excel);

        XSSFWorkbook workbook = new XSSFWorkbook(xlsx_in);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int lastRowIndex = sheet.getLastRowNum();

        for (int i = 1; i <= lastRowIndex; i++) {
            XSSFRow row = sheet.getRow(i);
            if (row == null) {
                break;
            }

            if (row.getCell(1) != null && row.getCell(1).getCellTypeEnum() == STRING) {
                String value = "";
                if (row.getCell(0) == null) {
                    continue;
                }
                if (row.getCell(0).getCellTypeEnum() == NUMERIC) {
                    value += row.getCell(0).getNumericCellValue();
                } else if (row.getCell(0).getCellTypeEnum() == STRING) {
                    value = row.getCell(0).getStringCellValue();
                }
                key_content_map.put(row.getCell(1).getStringCellValue().toLowerCase(), value);
            }

        }
        xlsx_in.close();
        return key_content_map;
    }

    private int getPictureFormat(String imgFile0) {
        int format;
        String imgFile = imgFile0.toLowerCase();
        if (imgFile.endsWith(".emf")) {
            format = XWPFDocument.PICTURE_TYPE_EMF;
        } else if (imgFile.endsWith(".wmf")) {
            format = XWPFDocument.PICTURE_TYPE_WMF;
        } else if (imgFile.endsWith(".pict")) {
            format = XWPFDocument.PICTURE_TYPE_PICT;
        } else if (imgFile.endsWith(".jpeg") || imgFile.endsWith(".jpg")) {
            format = XWPFDocument.PICTURE_TYPE_JPEG;
        } else if (imgFile.endsWith(".png")) {
            format = XWPFDocument.PICTURE_TYPE_PNG;
        } else if (imgFile.endsWith(".dib")) {
            format = XWPFDocument.PICTURE_TYPE_DIB;
        } else if (imgFile.endsWith(".gif")) {
            format = XWPFDocument.PICTURE_TYPE_GIF;
        } else if (imgFile.endsWith(".tiff")) {
            format = XWPFDocument.PICTURE_TYPE_TIFF;
        } else if (imgFile.endsWith(".eps")) {
            format = XWPFDocument.PICTURE_TYPE_EPS;
        } else if (imgFile.endsWith(".bmp")) {
            format = XWPFDocument.PICTURE_TYPE_BMP;
        } else if (imgFile.endsWith(".wpg")) {
            format = XWPFDocument.PICTURE_TYPE_WPG;
        } else {
            format = 0;
        }
        return format;
    }

    private void replaceFigure(XWPFParagraph paragraph, File image_f, int width, int height) throws IOException, InvalidFormatException {
        String imgFile1 = image_f.getName();

        FileInputStream pict = new FileInputStream(image_f);
        XWPFRun runn = paragraph.createRun();
        paragraph.setAlignment(ParagraphAlignment.CENTER);

        int format = getPictureFormat(imgFile1);
        int width_d, height_d;
        if (width > 0) {
            width_d = width;
        } else {
            width_d = 400;
        }
        if (height > 0) {
            height_d = height;
        } else {
            height_d = 200;
        }

        runn.addPicture(pict, format, imgFile1, Units.toEMU(width_d), Units.toEMU(height_d));
        pict.close();
    }

    private void process_TBL(String table_file_path) throws Exception {
        String value = table_file_path;
        boolean needCov = false;
        if (value.endsWith("doc")) {
            needCov = true;
            String temp_f = value.replace(".doc", ".docx");
            Document doc_x = new Document(value);
            doc_x.save(temp_f);
            value = temp_f;
        }

        File table_f = new File(value);
        FileInputStream table_file = new FileInputStream(table_f);

        XWPFDocument doc2 = new XWPFDocument(table_file);
        XWPFTable table_org = doc2.getTables().get(0);
        XWPFTable table_new = paragraph.getBody().insertNewTbl(cursor);

        table_new.removeRow(0);
        for (XWPFTableRow row_org : table_org.getRows()) {
            table_new.addRow(row_org);
        }
        table_file.close();
        if (needCov) {
            new File(value).delete();
        }
    }

    private void process_ctrs(List<CTR> ctrs) throws IOException, InvalidFormatException, Exception {
        boolean found = false;
        String content = "";
        for (CTR run : ctrs) {
            for (CTText text : run.getTList()) {
                if (text.getStringValue().indexOf("{") > -1) {
                    found = true;
                }

                if (found) {
                    content = content + text.getStringValue();
                    if ((text.getStringValue().indexOf("}")) > -1) {
                        process_content(text, content);
                        found = false;
                        content = "";
                    } else {
                        text.setStringValue("");
                    }
                }
            }
        }
    }

    private void process_Dictionary(JSONObject obj) throws IOException, InvalidFormatException {
        if (key_content_map == null) {
            key_content_map = getKey_value(new File(obj.getString("Dictionary")));
        }
    }

    private String process_TXT(JSONObject obj, String content) {
        String key = obj.getString("TXT");
        String val = key_content_map.get(key.toLowerCase());
        String newStr = "";

        if (val != null) {
            if (isPreview) {
                replace_values.add(val);
                newStr = content.replaceFirst("\\{\\{[^\\{\\}]+\\}\\}", placehoder);
            } else {
                newStr = content.replaceFirst("\\{\\{[^\\{\\}]+\\}\\}", val);
            }
        }
        return newStr;
    }

    private void process_IMG(XWPFParagraph paragraph, JSONObject obj) throws IOException, InvalidFormatException {
        File image_f = new File(obj.getString("IMG"));
        int width = 0, height = 0;
        try {
            width = Integer.valueOf(obj.getString("width"));
            height = Integer.valueOf(obj.getString("height"));
        } catch (Exception e) {
        }
        replaceFigure(paragraph, image_f, width, height);
    }

    private String process_json(String json, String content, String newStr) throws IOException, InvalidFormatException, Exception {
        JSONObject obj = JSONObject.fromObject(json);
        String str = newStr;
        if (json.indexOf("Dictionary:") > -1) {
            process_Dictionary(obj);
            str = "-1";
        } else if (json.indexOf("TXT:") > -1) {
            if (newStr.isEmpty()) {
                str = process_TXT(obj, content);
            } else {
                str = process_TXT(obj, newStr);
            }

        } else if (json.indexOf("TBL:") > -1) {
            process_TBL(obj.getString("TBL"));
            str = "-1";
        } else if (json.indexOf("IMG:") > -1) {
            process_IMG(paragraph, obj);
            str = "-1";
        }
        return str;
    }

    private void process_content(Object o, String content) throws InvalidFormatException, Exception {
        String newStr = "";
        Pattern p = Pattern.compile("\\{[^\\{\\}]+\\}");
        Matcher m = p.matcher(content);
        String json = "";
        replace_values.clear();
        while (m.find()) {
            json = m.group().replace("\\", "\\\\");
            newStr = process_json(json, content, newStr);
        }

        if (o instanceof XWPFRun) {
            XWPFRun xrun = (XWPFRun) o;
            if (newStr.equals("-1")) {
                xrun.setText("");
            } else {
                highlightReplacedtxt(paragraph, newStr, replace_values);
            }
        } else if (o instanceof CTText) {
            CTText text = (CTText) o;
            if (newStr.equals("-1")) {
                text.setStringValue("");
            } else {
                text.setStringValue(newStr);
            }
        }

    }

    private void process_ctrs2(List<CTR> new_ctrs) throws Exception {
        boolean found = false;
        String content = "";
        for (CTR run : new_ctrs) {
            if (isPreview && run.getRPr() != null) {
                run.unsetRPr();
            }
            XWPFRun xrun = new XWPFRun(run, paragraph);
            if (xrun.text().indexOf("{") > -1) {
                found = true;
            }

            if (found) {
                content = content + xrun.text();
                if (xrun.text().indexOf("}") > -1) {
                    process_content(xrun, content);
                    found = false;
                    content = "";
                }
            } else {
                XWPFRun r3 = paragraph.createRun();
                r3.setText(xrun.text());
            }

            for (CTText t : run.getTList()) {
                t.setStringValue("");
            }
        }

    }

    public void start(File input_word, File output_word, boolean isPre) throws FileNotFoundException, IOException, InvalidFormatException, Exception {
        isPreview = isPre;
        FileInputStream fin = new FileInputStream(input_word);
        XWPFDocument doc = new XWPFDocument(fin);
        cursor = doc.getDocument().getBody().newCursor();
        cursor.selectPath("./*");
        while (cursor.toNextSelection()) {
            XmlObject o = cursor.getObject();
            if (o instanceof CTP) {
                paragraph = new XWPFParagraph((CTP) o, doc);
                List<CTR> ctrs = paragraph.getCTP().getRList();
                List<CTR> new_ctrs = new ArrayList<>();
                for (CTR runx : ctrs) {
                    new_ctrs.add(runx);
                }
                if (isPreview) {
                    process_ctrs2(new_ctrs);
                } else {
                    process_ctrs(ctrs);
                }
            }
        }
        cursor.dispose();

        FileOutputStream fout = new FileOutputStream(output_word);
        doc.write(fout);
        fin.close();
        fout.close();
    }

    private void highlightReplacedtxt(XWPFParagraph paragraph, String newStr, List<String> vals) {
        String str = newStr;
        int beg, end;
        if (vals.size() < 1) {
            return;
        }
        for (String val : vals) {
            beg = str.indexOf(placehoder);
            end = beg + placehoder.length();
            if (beg > 0) {
                XWPFRun r1 = paragraph.createRun();
                r1.setText(str.substring(0, beg));
            }
            XWPFRun r2 = paragraph.createRun();
            r2.setText(val);
            CTShd cTShd = r2.getCTR().addNewRPr().addNewShd();
            cTShd.setFill("00FFFF");

            if (end < (str.length() - 1)) {
                str = str.substring(end + 1);
            }
        }
        int index_tail = newStr.lastIndexOf(placehoder) + placehoder.length();
        if (index_tail < (newStr.length() - 1)) {
            XWPFRun r3 = paragraph.createRun();
            r3.setText(newStr.substring(index_tail));
        }
    }

    /**
     * @param args the command line arguments
     */
    public static void run(String input_word_path, String output_word_path) throws InvalidFormatException, Exception {
        JavaApplication1 app = new JavaApplication1();
        //模板word文件路径
        File input_word = new File(input_word_path);
        //输出word文件路径
        File output_word = new File(output_word_path);

        String out_filename = output_word.getName();

        String name = out_filename.substring(0, out_filename.lastIndexOf("."));
        File output_word_pre = new File(output_word.getParent() + "\\" + name + "_pre.docx");
        app.start(input_word, output_word_pre, true);
        app.start(input_word, output_word, false);
    }

    public static void main(String[] args) throws IOException, FileNotFoundException, InvalidFormatException, Exception {
        JavaApplication1.run("./test_file/docMergeTpl.docx",
                "./test_file/my_out.docx");
         Document doc_x = new Document("G:\\NetBeansProjects\\JavaApplication1\\test_file\\docMergeTpl.docx");
            doc_x.save("G:\\NetBeansProjects\\JavaApplication1\\test_file\\docMergeTpl.pdf");

    }
}
