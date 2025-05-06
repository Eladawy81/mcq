import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.regex.*;
import java.util.*;

public class ImprovedMCQExtractor {

    static class MCQ {
        private String question;
        private Map<String, String> options = new LinkedHashMap<>();
        private String correctAnswer;
        private String explanation;

        // Getters & Setters
    }

    public static void main(String[] args) {
        String pdfPath = "TTTT.pdf";
        String excelPath = "MCQs_Improved.xlsx";

        try (PDDocument document = PDDocument.load(new File(pdfPath));
            Workbook workbook = new XSSFWorkbook();
            FileOutputStream out = new FileOutputStream(excelPath)) {

            PDFTextStripper stripper = new PDFTextStripper();
            stripper.setSortByPosition(true);
            String text = stripper.getText(document);

            Pattern pattern = Pattern.compile(
                "(\\d+)\\.\\s+((?:.*\\n)*?.*?)\\n" +
                "A\\.\\s+(.*?)\\n" +
                "B\\.\\s+(.*?)\\n" +
                "(?:C\\.\\s+(.*?)\\n)?" +
                "(?:D\\.\\s+(.*?)\\n)?" +
                "(?:E\\.\\s+(.*?)\\n)?" +
                "(?:.*?[Cc]orrect [Aa]nswer:\\s+([A-Ea-e]))" +
                "(?:\\n{2,}((?:.*\\n)*?.*?))?" +
                "(?=\\n\\d+\\.|\\z)",
                Pattern.DOTALL | Pattern.CASE_INSENSITIVE
            );

            List<MCQ> mcqList = new ArrayList<>();
            Matcher matcher = pattern.matcher(text);

            while (matcher.find()) {
                MCQ mcq = new MCQ();
                mcq.setQuestion(matcher.group(2).replaceAll("\\s+", " ").trim());
                mcq.getOptions().put("A", matcher.group(3).trim());
                mcq.getOptions().put("B", matcher.group(4).trim());
                if (matcher.group(5) != null) mcq.getOptions().put("C", matcher.group(5).trim());
                if (matcher.group(6) != null) mcq.getOptions().put("D", matcher.group(6).trim());
                if (matcher.group(7) != null) mcq.getOptions().put("E", matcher.group(7).trim());
                mcq.setCorrectAnswer(matcher.group(8).toUpperCase());
                if (matcher.group(9) != null) mcq.setExplanation(matcher.group(9).trim());
                mcqList.add(mcq);
            }

            Sheet sheet = workbook.createSheet("MCQs");
            createHeader(sheet);

            int rowIndex = 1;
            for (MCQ mcq : mcqList) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(mcq.getQuestion());
                row.createCell(1).setCellValue(mcq.getOptions().get("A"));
                row.createCell(2).setCellValue(mcq.getOptions().get("B"));
                row.createCell(3).setCellValue(mcq.getOptions().getOrDefault("C", ""));
                row.createCell(4).setCellValue(mcq.getOptions().getOrDefault("D", ""));
                row.createCell(5).setCellValue(mcq.getOptions().getOrDefault("E", ""));
                row.createCell(6).setCellValue(mcq.getCorrectAnswer());
                row.createCell(7).setCellValue(mcq.getExplanation());
            }

            workbook.write(out);
            System.out.println("✅ تم الاستخراج بنجاح! عدد الأسئلة: " + mcqList.size());

        } catch (IOException e) {
            System.err.println("خطأ في الملف: " + e.getMessage());
        } catch (PatternSyntaxException e) {
            System.err.println("خطأ في النمط: " + e.getMessage());
        } catch (Exception e) {
            System.err.println("خطأ غير متوقع: " + e.getMessage());
        }
    }

    private static void createHeader(Sheet sheet) {
        Row header = sheet.createRow(0);
        String[] columns = {"Question", "A", "B", "C", "D", "E", "Correct Answer", "Explanation"};
        for (int i = 0; i < columns.length; i++) {
            header.createCell(i).setCellValue(columns[i]);
        }
    }
}