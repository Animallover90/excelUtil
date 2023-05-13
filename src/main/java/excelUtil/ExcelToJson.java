package excelUtil;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;
import org.springframework.stereotype.Service;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

@Service
public class ExcelToJson {

    private final String exelFile = "C:/엑셀 소스 저장 위치/lotto.xlsx";

    public void exportJsonFileInBase10() throws IOException, InvalidFormatException {
        File file = new File(exelFile);

        XSSFWorkbook workbook = new XSSFWorkbook(file);

        // 해당 sheet를 이름으로 가져오기(엑셀 기본값 Sheet1)
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        // 해당 sheet의 데이터가 세로로 몇 줄인지 가져와서 반복문으로 값을 print
        int rowCnt = sheet.getPhysicalNumberOfRows();
        StringBuilder lastLottoNum = null;
        for (int i = 0; i < rowCnt; i++) {
            // 해당 row의(세로줄) 가로가 몇 줄인지 가져와서 또다시 반복문
            XSSFRow row = sheet.getRow(i);
            int colCnt = row.getPhysicalNumberOfCells();

            JSONObject obj = new JSONObject();
            StringBuilder lottoNum = new StringBuilder();
            for (int j = 0; j < colCnt; j++) {
                // 가로 cell의 값을 하나씩 print
                XSSFCell cell = row.getCell(j);
                lottoNum.append((int) Double.parseDouble(cell.toString()));
                if (j != colCnt-1) {
                    lottoNum.append(", ");
                }
            }
            obj.put("completion", lottoNum.toString());
            obj.put("prompt", "The last answer is " + lastLottoNum + ", Predict non-duplicate 6 from 1 to 45 numbers");

            if (i == 0) {
                obj.put("prompt", "Predict non-duplicate 6 from 1 to 45 numbers");
            }

            lastLottoNum = lottoNum;

            try {
                // 한줄씩 파일에 작성
                BufferedWriter jsonFile = new BufferedWriter(new FileWriter("C:/결과물 저장 위치/lotto.json", true));
                jsonFile.write(obj.toJSONString());
                jsonFile.newLine();
                jsonFile.flush();
                jsonFile.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public void exportJsonFileInBinary() throws IOException, InvalidFormatException {
        File file = new File(exelFile);

        XSSFWorkbook workbook = new XSSFWorkbook(file);

        // 해당 sheet를 이름으로 가져오기(엑셀 기본값 Sheet1)
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        // 해당 sheet의 데이터가 세로로 몇 줄인지 가져와서 반복문으로 값을 print
        int rowCnt = sheet.getPhysicalNumberOfRows();
        StringBuilder lastLottoNum = null;
        for (int i = 0; i < rowCnt; i++) {
            // 해당 row의(세로줄) 가로가 몇 줄인지 가져와서 또다시 반복문
            XSSFRow row = sheet.getRow(i);
            int colCnt = row.getPhysicalNumberOfCells();

            JSONObject obj = new JSONObject();
            StringBuilder lottoNum = new StringBuilder();
            for (int j = 0; j < colCnt; j++) {
                // 가로 cell의 값을 하나씩 print
                XSSFCell cell = row.getCell(j);
                lottoNum.append(Integer.toBinaryString((int) Double.parseDouble(cell.toString())));
                if (j != colCnt-1) {
                    lottoNum.append(", ");
                }
            }
            obj.put("completion", lottoNum.toString());
            obj.put("prompt", "The last answer is " + lastLottoNum + ", Predict non-duplicate 6 numbers from 1 to 45 in binary");

            if (i == 0) {
                obj.put("prompt", "Predict non-duplicate 6 from 1 to 45 numbers");
            }

            lastLottoNum = lottoNum;

            try {
                // 한줄씩 파일에 작성
                BufferedWriter jsonFile = new BufferedWriter(new FileWriter("C:/결과물 저장 위치/lottoBinary.json", true));
                jsonFile.write(obj.toJSONString());
                jsonFile.newLine();
                jsonFile.flush();
                jsonFile.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public void exportCSVFileIncludeCombined() throws IOException, InvalidFormatException {
        String savePath = "C:/결과물 저장 위치/lotto_combined.csv";

        try {
            // 한줄씩 파일에 작성
            BufferedWriter jsonFile = new BufferedWriter(new FileWriter(savePath, true));
            jsonFile.write("Num1,Num2,Num3,Num4,Num5,Num6,combined");
            jsonFile.newLine();
            jsonFile.flush();
            jsonFile.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        File file = new File(exelFile);

        XSSFWorkbook workbook = new XSSFWorkbook(file);

        // 해당 sheet를 이름으로 가져오기(엑셀 기본값 Sheet1)
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        // 해당 sheet의 데이터가 세로로 몇 줄인지 가져와서 반복문으로 값을 print
        int rowCnt = sheet.getPhysicalNumberOfRows();
        for (int i = 0; i < rowCnt; i++) {
            // 해당 row의(세로줄) 가로가 몇 줄인지 가져와서 또다시 반복문
            XSSFRow row = sheet.getRow(i);
            int colCnt = row.getPhysicalNumberOfCells();

            StringBuilder lottoNum = new StringBuilder();
            StringBuilder combined = new StringBuilder();
            for (int j = 0; j < colCnt; j++) {
                // 가로 cell의 값을 하나씩 print
                XSSFCell cell = row.getCell(j);
                lottoNum.append((int) Double.parseDouble(cell.toString()));
                lottoNum.append(",");

                combined.append("Num" + (j+1) + ":" + (int) Double.parseDouble(cell.toString()) + ";");

                if (j == colCnt-1) {
                    lottoNum.append(combined);
                }
            }

            try {
                // 한줄씩 파일에 작성
                BufferedWriter jsonFile = new BufferedWriter(new FileWriter(savePath, true));
                jsonFile.write(lottoNum.toString());
                jsonFile.newLine();
                jsonFile.flush();
                jsonFile.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
