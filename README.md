@Service
public class ExcelService {

    @Autowired
    private EntityARepository entityARepository;

    @Autowired
    private EntityBRepository entityBRepository;

    public void processExcelFile(MultipartFile file) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (isRowEmpty(row)) continue;

                EntityA entityA = new EntityA();
                EntityB entityB = new EntityB();

                // First 5 columns → Entity A
                entityA.setCol1(getCellValue(row.getCell(0)));
                entityA.setCol2(getCellValue(row.getCell(1)));
                entityA.setCol3(getCellValue(row.getCell(2)));
                entityA.setCol4(getCellValue(row.getCell(3)));
                entityA.setCol5(getCellValue(row.getCell(4)));

                // Next 5 columns → Entity B
                entityB.setCol6(getCellValue(row.getCell(5)));
                entityB.setCol7(getCellValue(row.getCell(6)));
                entityB.setCol8(getCellValue(row.getCell(7)));
                entityB.setCol9(getCellValue(row.getCell(8)));
                entityB.setCol10(getCellValue(row.getCell(9)));

                entityARepository.save(entityA);
                entityBRepository.save(entityB);
            }
        }
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return null;
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> null;
        };
    }

    private boolean isRowEmpty(Row row) {
        if (row == null) return true;
        for (int i = 0; i < 10; i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK &&
                (cell.getCellType() != CellType.STRING || !cell.getStringCellValue().trim().isEmpty())) {
                return false;
            }
        }
        return true;
    }
}
