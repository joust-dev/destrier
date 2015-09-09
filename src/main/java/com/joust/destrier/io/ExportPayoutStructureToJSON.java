package com.joust.destrier.io;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.common.collect.Maps;
import com.joust.game.contest.prize.PayoutStructure;
import com.joust.game.contest.prize.PointPrize;
import com.joust.game.contest.prize.PointPrizeRange;
import lombok.Data;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Map;

/**
 * Created by jreynoso on 9/8/15.
 * Joust, Inc.
 * <p/>
 * Read in first sheet of an XLSX file containing payout structure and export to JSON format.
 * Header row should be the winner ranks.
 * Subsequent rows should be the entry ranges with the respective payout percentage per winner rank.
 * e.g.
 * +---------------+-----+-----+
 * | Range \ Ranks |  1  |  2  |
 * +---------------+-----+-----+
 * |       2       | 1.0 |     |
 * +---------------+-----+-----+
 * |      3-10     | 0.7 | 0.3 |
 * +---------------+-----+-----+
 * Produces JSON in the following format:
 * {"pointPrizeRanges":[
 *   {"minEntries":2,"maxEntries":2,"prizes":[{"minRank":1,"maxRank":1,"prizePercent":100.00}]},
 *   {"minEntries":3,"maxEntries":10,"prizes":[{"minRank":1,"maxRank":1,"prizePercent":70.00},{"minRank":2,"maxRank":2,"prizePercent":30.00}]}
 * ]}
 * This format can be read into a com.joust.game.contest.prize.PayoutStructure POJO.
 **/

public class ExportPayoutStructureToJSON {

    private static final Logger log = LoggerFactory.getLogger(ExportPayoutStructureToJSON.class);

    private static final int RANK_ROW_INDEX = 0;    // index of row containing winner ranks
    private static final int RANGE_COL_INDEX = 0;   // index of column containing entry ranges

    private File inFile;
    private NumberFormat numberFormatter = NumberFormat.getInstance();
    private DecimalFormat decimalFormatter = new DecimalFormat("0.####");

    public ExportPayoutStructureToJSON(String in) {
        this.inFile = new File(in);
    }

    public void export(String out) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(inFile));

        PayoutStructure payoutStructure = new PayoutStructure();
        ArrayList<PointPrizeRange> ranges = new ArrayList<PointPrizeRange>();
        Map<Integer, MinMax> ranks = Maps.newHashMap();

        XSSFSheet sheet = wb.getSheetAt(0);
        int rows = sheet.getPhysicalNumberOfRows();
        log.info(String.format("\"%s\" has %d row(s)", wb.getSheetName(0), rows));
        for (int r = 0; r < rows; r++) {
            try {
                XSSFRow row = sheet.getRow(r);
                if (row == null) {
                    break;
                } else if (r == RANK_ROW_INDEX) {
                    ranks = parseRanks(row);
                } else {
                    PointPrizeRange range = parseRange(row, ranks);
                    if (range != null && !range.getPrizes().isEmpty()) {
                        ranges.add(range);
                    }
                }
            } catch (ParseException e) {
                log.error(String.format("Failed to parse row %d", r), e);
            }
        }

        payoutStructure.setPointPrizeRanges(ranges);
        exportToJSON(payoutStructure, new File(out));
    }

    private Map<Integer, MinMax> parseRanks(XSSFRow row) throws ParseException {
        Map<Integer, MinMax> ranks = Maps.newHashMap();

        int cells = row.getPhysicalNumberOfCells();
        for (int c = 0; c < cells; c++) {
            String stringVal = getCellValue(row.getCell(c));

            if (stringVal == null) {
                break;
            } else if (c != RANGE_COL_INDEX) {
                ranks.put(c, parseMinMax(stringVal));
            }
        }

        return ranks;
    }

    private PointPrizeRange parseRange(XSSFRow row, Map<Integer, MinMax> ranks) throws ParseException {
        PointPrizeRange range = new PointPrizeRange();
        ArrayList<PointPrize> prizes = new ArrayList<PointPrize>();

        int cells = row.getPhysicalNumberOfCells();
        for (int c = 0; c < cells; c++) {
            String stringVal = getCellValue(row.getCell(c));

            if (stringVal == null) {
                break;
            } else if (c == RANGE_COL_INDEX) {
                MinMax minMax = parseMinMax(stringVal);
                range.setMinEntries(minMax.getMin());
                range.setMaxEntries(minMax.getMax());
            } else {
                PointPrize prize = new PointPrize();
                MinMax rank = ranks.get(c);
                prize.setMinRank(rank.getMin());
                prize.setMaxRank(rank.getMax());
                BigDecimal val = new BigDecimal(stringVal);
                val = val.multiply(new BigDecimal(100));
                prize.setPrizePercent(val.setScale(2, BigDecimal.ROUND_DOWN));
                prizes.add(prize);
            }
        }
        range.setPrizes(prizes);

        return range;
    }

    private String getCellValue(XSSFCell cell) {
        String stringVal = null;
        Double numericVal = null;

        switch (cell.getCellType()) {
            case XSSFCell.CELL_TYPE_FORMULA:
            case XSSFCell.CELL_TYPE_NUMERIC:
                numericVal = cell.getNumericCellValue();
                break;
            case XSSFCell.CELL_TYPE_STRING:
                stringVal = cell.getStringCellValue();
                break;
            default:
        }

        if (stringVal == null && numericVal != null) {
            stringVal = decimalFormatter.format(numericVal);
        }

        return stringVal;
    }

    private MinMax parseMinMax(String val) throws ParseException {
        MinMax minMax = new MinMax();

        int separatorIdx = val.indexOf("-");
        if (separatorIdx > 0) {
            minMax.setMin(numberFormatter.parse(val.substring(0, separatorIdx)).longValue());
            minMax.setMax(numberFormatter.parse(val.substring(separatorIdx + 1)).longValue());
        } else {
            separatorIdx = val.indexOf("+");
            if (separatorIdx > 0) {
                minMax.setMin(numberFormatter.parse(val).longValue());
                minMax.setMax(null); // NOTE: if null, indicates no maximum
            } else {
                minMax.setVal(numberFormatter.parse(val).longValue());
            }
        }

        return minMax;
    }

    private void exportToJSON(PayoutStructure payoutStructure, File outFile) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        mapper.writeValue(outFile, payoutStructure);
    }

    public static void main(String[] args) {
        String in = args[0];
        String out = args[1];

        ExportPayoutStructureToJSON exportToJSON = new ExportPayoutStructureToJSON(in);
        try {
            exportToJSON.export(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Data
    private class MinMax {
        private Long min;
        private Long max;

        private void setVal(Long val) {
            setMin(val);
            setMax(val);
        }
    }
}
