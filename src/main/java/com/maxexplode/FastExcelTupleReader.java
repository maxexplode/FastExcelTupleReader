package com.maxexplode;

import com.maxexplode.format.BaseFormat;
import com.maxexplode.format.date.ExcelCustomDateFormat;
import com.maxexplode.stereotype.ExcelCellName;
import com.maxexplode.stereotype.ExcelRow;
import lombok.Getter;
import lombok.Setter;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Spliterator;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Consumer;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * This utility can read a very large Excel file in text streaming format with formulas included.
 * Compared to Poiji excel load this is way faster because of inlining with the sax event streaming.
 * This won't read any excel pivot tables for now, in future can enhance the same to read styles and pivots.
 *
 * @param <T>
 */
@Slf4j
public class FastExcelTupleReader<T> implements Spliterator<T> {

    /**
     * Idea how this works
     * <a href="https://en.wikipedia.org/wiki/Office_Open_XML">Office Open XML</a>
     * Excel format .xlsx is an openxml format (OOXML) developed by Microsoft, ECMA which is later extended with new schema features.
     * This is a Zipped XML-based file format. It includes xml files as resources if you unzip it. some of the main files it includes is
     * <ul>
     * <li>workbook.xml - Metadata related to excel</li>
     * <li>sharedStrings.xml - All the strings in Excel sheets are referred from here</li>
     * <li>xl/sheet1.xml - Each sheet in the Excel has its own xml file.</li>
     * </ul>
     * First we read all the string from the Excel sheet and build a cache out of it to refer later
     * When reading the sheet.xml.
     * After this we need to read the header from any given sheet and create a Stream out of rest of the
     * Rows remaining.
     */
    private final List<String> textCache = new ArrayList<>();
    private Map<Character, String> headerMap;
    private final String EXCEL_ROW = "excelRow";
    private Map<String, SetterInvocation> setterMap;
    private final String excelFilePath;
    private final ReadOptions readOptions;
    private final Class<T> targetClazz;
    private final AtomicInteger ROW_COUNT = new AtomicInteger(0);
    private XMLStreamReader rowParser = null;
    private ZipFile zipFile;
    private final List<BaseFormat> registeredFormatters = new ArrayList<>();
    private final List<Integer> nmFmtMap = new ArrayList<>();
    private final Map<Integer, BaseFormat> baseFormaterMap = new HashMap<>();
    private boolean dataRow = false;

    public FastExcelTupleReader(String excelFilePath, Class<T> targetClazz, ReadOptions readOptions) throws XMLStreamException, IOException {
        this.excelFilePath = excelFilePath;
        this.targetClazz = targetClazz;
        this.readOptions = readOptions;
        initializeWorkBook();
    }

    /**
     * Initializing the necessary objects and caches to allow parsing and reading data from the specified sheet in the Excel workbook file.
     *
     * @throws IOException
     * @throws XMLStreamException
     */
    private void initializeWorkBook() throws IOException, XMLStreamException {
        Path filePath = Paths.get(excelFilePath);

        if (!Files.exists(filePath)) {
            throw new IOException("File does not exists : " + excelFilePath);
        }

        registerFormatters(new ExcelCustomDateFormat());

        zipFile = new ZipFile(excelFilePath);

        setterMap = Collections.unmodifiableMap(buildSetterMap());
        XMLInputFactory factory = XMLInputFactory.newInstance();
        buildStringCache(factory);
        buildStylesCache(factory);
        ZipEntry sheet = zipFile.getEntry(String.format("xl/worksheets/sheet%s.xml", readOptions.getSheetIdx()));
        InputStream inSheet = zipFile.getInputStream(sheet);

        rowParser = factory.createXMLStreamReader(inSheet);
        //Header
        readRow();
    }

    private void registerFormatters(BaseFormat format) {
        registeredFormatters.add(format);
        for (Integer formatId : format.supportedFormats()) {
            baseFormaterMap.put(formatId, format);
        }
    }

    public Stream<T> read() {
        return buildStream();
    }

    /**
     * This is parsing the "sharedStrings.xml" file from an Excel (.xlsx) file zip archive
     * to build a cache of all the strings used in the file.
     * It uses an XMLStreamReader to iteratively parse the XML file. When it encounters tags it sets a flag to expect text.
     * When it then encounters character data it adds that text to a textCache and unset the expectText flag.
     *
     * @throws IOException        In of any Disk I/O issue
     * @throws XMLStreamException If Stream is interrupted or an error occur within
     * @see XMLInputFactory
     */
    private void buildStringCache(XMLInputFactory factory) throws IOException, XMLStreamException {
        long start = System.currentTimeMillis();
        //Parse String files
        ZipEntry sharedStrings = zipFile.getEntry("xl/sharedStrings.xml");
        InputStream inputStream = zipFile.getInputStream(sharedStrings);
        XMLStreamReader parser = factory.createXMLStreamReader(inputStream);
        try {
            boolean expectText = false;
            while (true) {
                int next = parser.next();

                if (next == XMLStreamConstants.END_DOCUMENT) {
                    parser.close();
                    break;
                }

                if (next == XMLStreamConstants.START_ELEMENT) {
                    String localName = parser.getLocalName();
                    switch (localName) {
                        /*
                        - "sst" = Root element, contains the entire shared string table
                        - "si" = Individual shared string item within the table
                         */
                        case "sst", "si" -> {
                            continue;
                        }
                        /*
                            Value it holds
                         */
                        case "t" -> expectText = true;
                    }

                }

                if (expectText && next == XMLStreamConstants.CHARACTERS) {
                    textCache.add(parser.getText());
                    expectText = false;
                }
            }
        } finally {
            if (parser != null) {
                parser.close();
            }
        }

        log.atInfo().setMessage("Text cache processed {} in {}s")
                .addArgument(textCache::size)
                .addArgument(() -> (System.currentTimeMillis() - start) / 1000.0).log();

    }

    /**
     * This is parsing the "styles.xml" file from an Excel (.xlsx) file zip archive
     * It uses an XMLStreamReader to iteratively parse the XML file. When it encounters tags it sets a flag to expect text.
     * When it then encounters character data it adds that text to a textCache and unset the expectText flag.
     *
     * @throws IOException
     * @throws XMLStreamException
     */
    private void buildStylesCache(XMLInputFactory factory) throws IOException, XMLStreamException {
        long start = System.currentTimeMillis();
        //Parse style file
        ZipEntry styles = zipFile.getEntry("xl/styles.xml");
        InputStream inputStream = zipFile.getInputStream(styles);
        XMLStreamReader parser = factory.createXMLStreamReader(inputStream);

        Map<Integer, String> customFormats = new HashMap<>();

        try {
            boolean nmFmt = false;
            Integer nmFmtId = null;
            while (true) {
                int next = parser.next();

                if (next == XMLStreamConstants.END_DOCUMENT) {
                    parser.close();
                    break;
                } else if (next == XMLStreamConstants.START_ELEMENT) {
                    String localName = parser.getLocalName();
                    switch (localName) {
                        case "cellXfs" -> {
                            nmFmt = true;
                        }
                        case "xf" -> {
                            if (nmFmt) {
                                for (int i = 0; i < parser.getAttributeCount(); i++) {
                                    QName attributeName = parser.getAttributeName(i);
                                    if (attributeName.getLocalPart().equals("numFmtId")) {
                                        nmFmtId = Integer.parseInt(parser.getAttributeValue(i));
                                    }
                                }
                            }
                        }
                        case "numFmt" -> {
                            Integer numFmtId = null;
                            for (int i = 0; i < parser.getAttributeCount(); i++) {
                                QName attributeName = parser.getAttributeName(i);
                                if (attributeName.getLocalPart().equals("numFmtId")) {
                                    numFmtId = Integer.parseInt(parser.getAttributeValue(i));
                                } else if (attributeName.getLocalPart().equals("formatCode")) {
                                    customFormats.put(numFmtId, parser.getAttributeValue(i));
                                }
                            }
                        }
                    }
                } else if (next == XMLStreamConstants.END_ELEMENT) {
                    String localName = parser.getLocalName();
                    switch (localName) {
                        case "cellXfs" -> {
                            nmFmt = false;
                        }
                        case "xf" -> {
                            if (null == nmFmtId) {
                                continue;
                            }
                            String format = customFormats.get(nmFmtId);
                            BaseFormat defaultFormat = BaseFormat.any();
                            if (null != format && !baseFormaterMap.containsKey(nmFmtId)) {
                                for (BaseFormat baseFormat : registeredFormatters) {
                                    if (baseFormat.supports(nmFmtId, format)) {
                                        defaultFormat = baseFormat;
                                        break;
                                    }
                                }
                            }
                            baseFormaterMap.putIfAbsent(nmFmtId, defaultFormat);
                            nmFmtMap.add(nmFmtId);
                        }
                    }
                }
            }
        } finally {
            if (parser != null) {
                parser.close();
            }
        }

        log.atInfo().setMessage("Text cache processed {} in {}s")
                .addArgument(textCache::size)
                .addArgument(() -> (System.currentTimeMillis() - start) / 1000.0).log();

    }

    private Stream<T> buildStream() {
        //Creating a row stream :)
        return StreamSupport.stream(this, false)
                .onClose(this::close);
    }

    /**
     * This is reading data from an XML stream and processing it row by row.
     * It uses an XMLStreamReader to iterate over the XML elements. For each start element,
     * it checks the element name and extracts attributes like the column position and type.
     * When it encounters a "v" element, it sets a flag to read the character content for the column value.
     * The character content is read and stored in a variable.
     * For end elements, it processes the row or column value based on the element name.
     * Row elements trigger processing of the header or data row. Column elements add the value to the current row.
     * The processed rows are used to create an object of type T, which is returned at the end after the full XML stream is read.
     *
     * @return T
     * @throws XMLStreamException
     */
    private T readRow() throws XMLStreamException {
        boolean endRow = false;
        boolean afterVal = false;

        String colPos = null;
        String stylePos = null;
        String colType = null;
        Object currVal = null;
        boolean headerRow = true;

        ResultRow currentResult = null;
        T processObject = null;
        while (!endRow) {
            int next = rowParser.next();

            if (next == XMLStreamConstants.END_DOCUMENT) {
                rowParser.close();
                break;
            }

            switch (next) {
                case XMLStreamConstants.START_ELEMENT -> {
                    String localName = rowParser.getLocalName();
                    switch (localName) {
                        case "row" -> {
                            String rowPos = rowParser.getAttributeValue(0);
                            headerRow = rowPos.equals(readOptions.getHeaderRowIdx());
                            if (!dataRow) {
                                dataRow = rowPos.equals(readOptions.getDataRowIdx());
                            }
                            currentResult = new ResultRow();
                            currentResult.setRow(rowPos);
                        }
                        case "c" -> {
                            //Column index
                            colType = null;
                            stylePos = null;
                            for (int i = 0; i < rowParser.getAttributeCount(); i++) {
                                QName attributeName = rowParser.getAttributeName(i);
                                switch (attributeName.getLocalPart()) {
                                    case "t" -> colType = rowParser.getAttributeValue(i);
                                    case "r" -> colPos = rowParser.getAttributeValue(i);
                                    case "s" -> stylePos = rowParser.getAttributeValue(i);
                                }
                            }
                        }
                        case "v" -> {
                            afterVal = true;
                        }
                    }
                }
                case XMLStreamConstants.CHARACTERS -> {
                    if (afterVal) {
                        currVal = rowParser.getText();
                    }
                }
                case XMLStreamConstants.END_ELEMENT -> {
                    String localName = rowParser.getLocalName();
                    switch (localName) {
                        case "row" -> {
                            if (headerRow) {
                                processHeader(currentResult);
                            } else {
                                if (dataRow && !(readOptions.isSkipNull() && currentResult.isEmpty())) {
                                    ROW_COUNT.incrementAndGet();
                                    processObject = processObject(currentResult);
                                }
                            }
                            currentResult = null;
                            endRow = true;
                        }
                        case "c" -> {
                            char col = colPos.charAt(0);
                            if (!headerRow && null != stylePos) {
                                //found a formatted date
                                currentResult.add(col, findStyle(stylePos, colType, currVal));
                            } else if (null != colType && colType.equals("s")) {
                                currentResult.add(col, Integer.parseInt((String) currVal));
                            } else {
                                currentResult.add(col, currVal);
                            }
                        }
                        case "v" -> afterVal = false;
                    }
                }
            }
        }
        return processObject;
    }

    private Object findStyle(String stylePos, String colType, Object currVal) {
        if (currVal == null) {
            return null;
        }
        Integer nmFmtId = nmFmtMap.get(Integer.parseInt(stylePos));
        if (null == nmFmtId) {
            throw new RuntimeException("Unable to find style id : " + stylePos);
        }
        if (0 == nmFmtId && "s".equals(colType)) {
            //This is a general number format, which is string, so value should be coming from stringcache
            return Integer.parseInt((String) currVal);
        }
        BaseFormat baseFormat = baseFormaterMap.get(nmFmtId);
        if (null == baseFormat) {
            throw new RuntimeException("No valid formatter found for nmFmtId : " + nmFmtId);
        }
        return baseFormat.format(nmFmtId, (String) currVal);
    }

    private void processHeader(ResultRow headerRow) {
        Map<Character, String> temp = new HashMap<>();
        headerRow.getPositionsMap().forEach((colPos, valPos) -> {
            if (valPos instanceof Integer index) {
                temp.put(colPos, textCache.get(index));
            }
        });
        headerMap = Collections.unmodifiableMap(temp);
    }

    private T processObject(ResultRow resultRow) {
        try {
            T instance = targetClazz.getConstructor().newInstance();
            resultRow.getPositionsMap().forEach((colPos, value) -> {
                SetterInvocation invocation = setterMap.get(headerMap.get(colPos));
                //No need to map any invocation which is not having a header
                if (null != invocation) {
                    invocation.invoke(targetClazz, instance, value);
                }
            });
            if (setterMap.containsKey(EXCEL_ROW)) {
                SetterInvocation excelRowInvocation = setterMap.get(EXCEL_ROW);
                excelRowInvocation.invoke(targetClazz, instance, resultRow.getRow());
            }
            return instance;

        } catch (InstantiationException | IllegalAccessException | InvocationTargetException |
                 NoSuchMethodException e) {
            throw new RuntimeException(e);
        }
    }

    private Map<String, SetterInvocation> buildSetterMap() {
        Map<String, SetterInvocation> setterMap = new HashMap<>();
        for (Field field : targetClazz.getDeclaredFields()) {
            ExcelCellName annotation = field.getAnnotation(ExcelCellName.class);
            String fieldName = field.getName();
            if (annotation != null) {
                String value = annotation.value();
                if(StringUtils.isBlank(value)){
                    value = fieldName;
                }
                setterMap.put(value, new SetterInvocation(value, fieldName, field.getType()));
            }
            ExcelRow excelRow = field.getAnnotation(ExcelRow.class);
            if (excelRow != null) {
                setterMap.put(EXCEL_ROW, new SetterInvocation(EXCEL_ROW, fieldName, field.getType()));
            }
        }
        return setterMap;
    }

    @SneakyThrows
    @Override
    public boolean tryAdvance(Consumer<? super T> action) {
        T row = readRow();
        if (row != null) {
            action.accept(row);
            return true;
        }
        return false;
    }

    @Override
    public Spliterator<T> trySplit() {
        //Split is not yet supported
        return null;
    }

    @Override
    public long estimateSize() {
        //Goes along with split
        return 0;
    }

    @Override
    public int characteristics() {
        return Spliterator.ORDERED | Spliterator.NONNULL;
    }

    private void close() {
        if (rowParser != null) {
            try {
                rowParser.close();
            } catch (XMLStreamException e) {
                throw new RuntimeException(e);
            }
        }
        if (zipFile != null) {
            try {
                zipFile.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    @Getter
    private class SetterInvocation {

        private final String excelColumnName;
        private final String fieldName;
        private final Class<?> fieldType;

        public SetterInvocation(String excelColumnName, String fieldName, Class<?> fieldType) {
            this.excelColumnName = excelColumnName;
            this.fieldName = fieldName;
            this.fieldType = fieldType;
        }

        private String setter() {
            return "set" + Character.toUpperCase(fieldName.charAt(0)) + fieldName.substring(1);
        }

        public void invoke(Class<?> target, Object instance, Object value) {
            Method declaredMethod;
            try {
                declaredMethod = target.getDeclaredMethod(setter(), fieldType);
                if (value instanceof Integer index) {
                    declaredMethod.invoke(instance, textCache.get(index));
                } else {
                    if (fieldType.equals(Long.class)) {
                        declaredMethod.invoke(instance, Long.parseLong(value.toString()));
                    } else if (fieldType.equals(Integer.class)) {
                        declaredMethod.invoke(instance, Integer.parseInt(value.toString()));
                    } else if (fieldType.equals(Double.class)) {
                        declaredMethod.invoke(instance, Double.parseDouble(value.toString()));
                    } else if (fieldType.equals(String.class)) {
                        declaredMethod.invoke(instance, String.valueOf(value));
                    }
                }
            } catch (NoSuchMethodException | IllegalArgumentException |  IllegalAccessException | InvocationTargetException e) {
                throw new RuntimeException(e);
            }
        }
    }

    @Getter
    static private class ResultRow {
        @Setter
        private String row;
        private final Map<Character, Object> positionsMap = new HashMap<>();

        public void add(Character colPos, Object val) {
            if (val != null) {
                positionsMap.put(colPos, val);
            }
        }

        public boolean isEmpty() {
            return positionsMap.isEmpty();
        }
    }

    public Map<Character, String> headers() {
        return this.headerMap;
    }

    public int getTotalRowCount() {
        return ROW_COUNT.get();
    }
}