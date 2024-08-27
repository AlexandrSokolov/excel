package com.savdev.demo.excel.service.validators;

import com.savdev.commons.excel.dto.ExcelLine;
import com.savdev.demo.excel.service.validation.Validator;
import org.springframework.stereotype.Service;

import java.util.Map;
import java.util.Optional;

import static com.savdev.demo.excel.config.ExcelConstants.COLUMMN_A_ATTRIBUTE_NAME;

@Service
public class ColumnAValidator implements Validator<ExcelLine> {

  @Override
  public Optional<String> errorIfNotValid(ExcelLine excelLine) {
    var maybeValue = Optional.ofNullable(
      excelLine.getColumnName2Attribute2Value()
        .get("A"))
      .map(Map.Entry::getValue);

    if (maybeValue.isEmpty()) {
      return Optional.of("Column 'A -> " + COLUMMN_A_ATTRIBUTE_NAME + "' is required");
    }

    return Optional.empty();

  }
}
