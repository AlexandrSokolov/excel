package com.savdev.demo.excel.service.validation;

import java.util.Optional;

public interface Validator<T> {

  Optional<String> errorIfNotValid(T item2Validate);
}
