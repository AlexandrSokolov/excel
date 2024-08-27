package com.savdev.demo.excel.service.validators;

import com.savdev.commons.excel.dto.ExcelLine;
import com.savdev.demo.excel.service.validation.Validator;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.util.Map;
import java.util.Optional;

import static com.savdev.demo.excel.config.ExcelConstants.COLUMMN_E_ATTRIBUTE_NAME;

@Service
public class ColumnEValidator implements Validator<ExcelLine> {

  @Override
  public Optional<String> errorIfNotValid(ExcelLine excelLine) {
    var maybeValue = Optional.ofNullable(
        excelLine.getColumnName2Attribute2Value()
          .get("E"))
      .map(Map.Entry::getValue);

    if (maybeValue.isPresent()
      && ! (maybeValue.get() instanceof BigDecimal)) {
      return Optional.of(
        "Column 'E -> " + COLUMMN_E_ATTRIBUTE_NAME + "' is expected to be number");
    }

    return Optional.empty();
  }
}
