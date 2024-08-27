package com.savdev.demo.excel.service.validation.collector;

import java.util.List;

public record InvalidItem<T>(
  T originalItem,
  String itemRepresentation,
  List<String> validationErrors
) {
}
