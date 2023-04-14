import java.io.*;

import fr.opensagres.poi.xwpf.converter.core.ImageManager;
import fr.opensagres.poi.xwpf.converter.core.XWPFConverterException;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.*;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.fit.pdfdom.PDFDomTree;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class WordToHtmlConverterA {

    public static void main(String[] args) throws Throwable {
        //convertDoc("input.docx");
        //convertDocx("input.docx");
        //convertAddImgTag("input.docx");
        convertDocxToHtml("input.docx");
    }

    //docx -> pdf -> html
    private static void convertDocxToHtml(String fileName) throws IOException {
        //pdf
        String docPath = fileName;
        String pdfPath = "output.pdf";

        InputStream in = new FileInputStream(new File(docPath));
        XWPFDocument document = new XWPFDocument(in);
        PdfOptions options = PdfOptions.create();
        OutputStream out = new FileOutputStream(new File(pdfPath));
        PdfConverter.getInstance().convert(document, out, options);

        document.close();
        out.close();

        //html
        PDDocument pdf = PDDocument.load(new File(pdfPath));
        Writer output = new PrintWriter("output.html", "utf-8");
        new PDFDomTree().writeText(pdf, output);
        output.close();
        System.out.println("Done");
    }

    private static void convertAddImgTag(String fileName) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(new File(fileName));
        XWPFDocument document = new XWPFDocument(fileInputStream);

        XHTMLOptions options = XHTMLOptions.create();
        options.setImageManager(new ImageManager(new File("./"), "images"));

        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    if (cell.getText().trim().isEmpty()) {
                        cell.setText(" ");
                    }
                }
            }
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        XHTMLConverter.getInstance().convert(document, outputStream, options);
        String html = outputStream.toString();

        Document doc = Jsoup.parse(html);
        Elements imgs = doc.select("img");
        for (Element img : imgs) {
            img.attr("style", "float: left; margin-right: 10px;");
        }

        BufferedWriter writer = new BufferedWriter(new FileWriter("output.html"));
        writer.write(doc.outerHtml());
        writer.close();
        System.out.println("Done");
    }

    private static void convertDocx(String fileName) {
        try {
            String outputFilePath = "output.html";
            FileInputStream fis = new FileInputStream(fileName);
            XWPFDocument document = new XWPFDocument(fis);
            FileOutputStream fos = new FileOutputStream(outputFilePath);
            XHTMLOptions options = XHTMLOptions.create();
            options.setImageManager(new ImageManager(new File("./"), "images"));

            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        if (cell.getText().trim().isEmpty()) {
                            cell.setText(" ");
                        }
                    }
                }
            }

            XHTMLConverter.getInstance().convert(document, fos, options);

            fis.close();
            fos.close();

            System.out.println("Done");
        } catch (XWPFConverterException e) {
            System.out.println("Error");
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static void convertDoc(String fileName) {
        try {
            File file = new File(fileName);
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());

            HWPFDocument document = new HWPFDocument(fis);


            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                    DocumentBuilderFactory.newInstance()
                            .newDocumentBuilder().newDocument()
            );
            wordToHtmlConverter.setPicturesManager(new PicturesManager() {
                @Override
                public String savePicture(byte[] bytes, PictureType pictureType, String s, float v, float v1) {
                    // Сохраняем картинку в файл
                    File imageFile = new File(s);
                    try (FileOutputStream fos = new FileOutputStream(imageFile)) {
                        fos.write(bytes);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    // Возвращаем путь к сохраненному файлу
                    return imageFile.getAbsolutePath();
                }
            });
            wordToHtmlConverter.processDocument(document);

            org.w3c.dom.Document htmlDocument = wordToHtmlConverter.getDocument();
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(out);

            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);

            String html = new String(out.toByteArray());
            FileWriter writer = new FileWriter("output.html");
            writer.write(html);
            writer.close();
            System.out.println("Done");
        } catch (TransformerConfigurationException e) {
            throw new RuntimeException(e);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (ParserConfigurationException e) {
            throw new RuntimeException(e);
        } catch (TransformerException e) {
            throw new RuntimeException(e);
        }

    }
}





















//        String filename = "_заявка на Транссервис.doc";
//        File file = new File(filename);
//        FileInputStream fis = new FileInputStream(file.getAbsolutePath());
//
//        HWPFDocument document = new HWPFDocument(fis);
//
//
//        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
//                DocumentBuilderFactory.newInstance()
//                        .newDocumentBuilder().newDocument()
//        );
//        wordToHtmlConverter.setPicturesManager(new PicturesManager() {
//            @Override
//            public String savePicture(byte[] bytes, PictureType pictureType, String s, float v, float v1) {
//                // Сохраняем картинку в файл
//                File imageFile = new File(s);
//                try (FileOutputStream fos = new FileOutputStream(imageFile)) {
//                    fos.write(bytes);
//                } catch (IOException e) {
//                    e.printStackTrace();
//                }
//                // Возвращаем путь к сохраненному файлу
//                return imageFile.getAbsolutePath();
//            }
//        });
//        wordToHtmlConverter.processDocument(document);
//
//        org.w3c.dom.Document htmlDocument = wordToHtmlConverter.getDocument();
//        ByteArrayOutputStream out = new ByteArrayOutputStream();
//        DOMSource domSource = new DOMSource(htmlDocument);
//        StreamResult streamResult = new StreamResult(out);
//
//        TransformerFactory tf = TransformerFactory.newInstance();
//        Transformer serializer = tf.newTransformer();
//        serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
//        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
//        serializer.setOutputProperty(OutputKeys.METHOD, "html");
//        serializer.transform(domSource, streamResult);
//
//        String html = new String(out.toByteArray());
//        FileWriter writer = new FileWriter("output10.html");
//        writer.write(html);
//        writer.close();
//        System.out.println("Успех!");
//    }





















    //        String docPath = "заявка на Транссервис.docx";
//        String root = "./";
//        String htmlPath = "output3.html";
//
//        XWPFDocument document = new XWPFDocument(new FileInputStream(docPath));
//
//        XHTMLOptions options = XHTMLOptions.create().setImageManager(new ImageManager(new File(root), "images"));
//        FileOutputStream out = new FileOutputStream(htmlPath);
//        XHTMLConverter.getInstance().convert(document, out, options);
//
//        out.close();
////        document.close();
//}
//    }

//        XHTMLOptions options = XHTMLOptions.create().setImageManager(new ImageManager(new File(root), "images"));
//
//

//
//
//
////
//    }
//}





