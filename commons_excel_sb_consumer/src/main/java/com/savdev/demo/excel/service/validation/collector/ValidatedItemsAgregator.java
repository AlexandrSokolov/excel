package com.savdev.demo.excel.service.validation.collector;

import java.util.List;
import java.util.function.Function;

public record ValidatedItemsAgregator<T, R>(
  List<R> validItems,
  List<InvalidItem<T>> invalidItems,
  Function<T, R> validItemMapper) {
}
