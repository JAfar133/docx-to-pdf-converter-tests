package tech;

import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;


public class Test {
    public static void main(String[] args) throws Exception {
        // Укажите путь к установке LibreOffice
        String officeHome = "C:/Program Files/LibreOffice"; // Измените этот путь в зависимости от вашей системы

        // Создаем экземпляр менеджера офиса
        LocalOfficeManager officeManager = LocalOfficeManager.builder()
                .officeHome(new File(officeHome))
                .install().build();

        try {
            // Запускаем менеджер офиса
            officeManager.start();

            // Создаем экземпляр конвертера
            DocumentConverter converter = LocalConverter.make(officeManager);

            // Указываем входной DOCX файл и выходной PDF файл
            File inputFile = new File("./input.docx");
            File outputFile = new File("./output.pdf");

            // Выполняем конвертацию
            converter.convert(inputFile).to(outputFile).execute();

            System.out.println("Конвертация успешно завершена!");

        } catch (OfficeException e) {
            e.printStackTrace();
        } finally {
            // Останавливаем менеджер офиса
            try {
                officeManager.stop();
            } catch (OfficeException e) {
                e.printStackTrace();
            }
        }
    }
}
