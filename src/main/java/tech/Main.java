package tech;

import com.itextpdf.text.*;
import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.*;
import java.math.BigInteger;

public class Main {
    public static void main(String[] args) {
        try (FileInputStream docxInput = new FileInputStream("input.docx");
             FileOutputStream pdfOutput = new FileOutputStream("output.pdf")) {

            XWPFDocument docx = new XWPFDocument(docxInput);
            CTSectPr sectPr = docx.getDocument().getBody().getSectPr();
            CTPageMar pageMar = sectPr.getPgMar();
            float leftMargin = ((BigInteger) pageMar.getLeft()).intValue() * 0.05f;
            float rightMargin = ((BigInteger) pageMar.getRight()).intValue() * 0.05f;
            float topMargin = ((BigInteger) pageMar.getTop()).intValue() * 0.05f;
            float bottomMargin = ((BigInteger) pageMar.getBottom()).intValue() * 0.05f;
            Document pdf = new Document(PageSize.A4, leftMargin, rightMargin, topMargin, bottomMargin);

            int defaultFontSize = docx.getStyles().getDefaultRunStyle().getFontSizeAsDouble().intValue();

            PdfWriter.getInstance(pdf, pdfOutput);
            pdf.open();
            int style1 = Font.NORMAL;
            for (XWPFParagraph para : docx.getParagraphs()) {
                StringBuilder sb = new StringBuilder();
                if (para.getRuns().isEmpty()) {
                    pdf.add(new Paragraph(" "));
                    continue;
                }

                float indentLeft = para.getIndentationLeft() * 0.05f;
                float indentRight = para.getIndentationRight() * 0.05f;

                Font font = new Font(Font.FontFamily.TIMES_ROMAN, defaultFontSize, style1);
                for (XWPFRun run : para.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) {
                        int fontSize = run.getFontSize();
                        if (fontSize == -1) fontSize = defaultFontSize;

                        int style = Font.NORMAL;
                        if (run.isBold()) style |= Font.BOLD;
                        if (run.isItalic()) style |= Font.ITALIC;

                        font = new Font(Font.FontFamily.TIMES_ROMAN, fontSize, style);

                        sb.append(text);
                    }
                    if (run.getEmbeddedPictures().size() > 0) {
                        for (XWPFPicture picture : run.getEmbeddedPictures()) {
                            InputStream picStream = new ByteArrayInputStream(picture.getPictureData().getData());
                            Image img = Image.getInstance(picStream.readAllBytes());
                            pdf.add(img);
                        }
                    }
                }
                Paragraph p = new Paragraph(sb.toString(), font);
                int spacingAfter = para.getSpacingAfter();
                int spacingBefore = para.getSpacingBefore();
                double spacingBetween = para.getSpacingBetween();
                para.setSpacingBetween(spacingBetween);

                if (spacingAfter == -1) {
                    if (para.getSpacingLineRule().equals(LineSpacingRule.AUTO)) {
                        p.setSpacingAfter(2f);
                    }
                } else {
                    p.setSpacingAfter(spacingAfter);
                }
                if (spacingBefore == -1) {
                    if (para.getSpacingLineRule().equals(LineSpacingRule.AUTO)) {
                        p.setSpacingBefore(2f);
                    }
                } else {
                    p.setSpacingBefore(spacingBefore);
                }

                p.setIndentationLeft(indentLeft);
                p.setIndentationRight(indentRight);

                pdf.add(p);
            }

            pdf.close();
            docx.close();

            System.out.println("DOCX with images converted to PDF successfully.");
        } catch (IOException | DocumentException e) {
            e.printStackTrace();
        }
    }

}