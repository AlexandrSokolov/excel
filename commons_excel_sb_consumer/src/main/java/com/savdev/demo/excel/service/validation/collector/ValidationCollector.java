package com.savdev.demo.excel.service.validation.collector;

import com.savdev.demo.excel.service.validation.Validator;

import java.util.Collections;
import java.util.LinkedList;
import java.util.Optional;
import java.util.Set;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.BinaryOperator;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.stream.Collector;
import java.util.stream.Collectors;
import java.util.stream.Stream;


/**
 *
 * @param <R> result of mapping of valid items from T
 * @param <T> original input item
 *
 */
public class ValidationCollector<T, R>
  implements Collector<
  T,
  ValidatedItemsAgregator<T, R>,
  ValidatedItemsAgregator<T, R>> {

  private final List<Validator<T>> itemValidators;
  private final Function<T, R> validItemMapper;
  private final Function<T, String> invalidItemDescription;

  public ValidationCollector(
    List<Validator<T>> itemValidators,
    Function<T, R> validItemMapper,
    Function<T, String> invalidItemDescription) {
    this.itemValidators = itemValidators;
    this.validItemMapper = validItemMapper;
    this.invalidItemDescription = invalidItemDescription;
  }

  public ValidationCollector(
    List<Validator<T>> itemValidators,
    Function<T, R> validItemMapper) {
    this.itemValidators = itemValidators;
    this.validItemMapper = validItemMapper;
    this.invalidItemDescription = Object::toString;
  }


  @Override
  public Supplier<ValidatedItemsAgregator<T, R>> supplier() {
    return () -> new ValidatedItemsAgregator<>(
      new LinkedList<>(),
      new LinkedList<>(),
      validItemMapper );
  }

  @Override
  public BiConsumer<ValidatedItemsAgregator<T, R>, T> accumulator() {
    return (agregator, originalItem) -> {

      var validationErrors = itemValidators.stream()
        .map(validator -> validator.errorIfNotValid(originalItem))
        .filter(Optional::isPresent)
        .map(Optional::get)
        .toList();

      if (validationErrors.isEmpty()) {
        agregator.validItems()
          .add(
            validItemMapper.apply(originalItem));
      } else {
        agregator.invalidItems().add(
          new InvalidItem<>(
            originalItem,
            this.invalidItemDescription.apply(originalItem),
            validationErrors) );
      }
    };
  }

  @Override
  public BinaryOperator<ValidatedItemsAgregator<T, R>> combiner() {
    return (l, r) -> new ValidatedItemsAgregator<>(
      Stream.concat(
        l.validItems().stream(),
        r.validItems().stream())
        .toList(),
      Stream.concat(
          l.invalidItems().stream(),
          r.invalidItems().stream())
        .toList(),
      validItemMapper);
  }

  @Override
  public Function<ValidatedItemsAgregator<T, R>, ValidatedItemsAgregator<T, R>> finisher() {
    return Function.identity();
  }

  @Override
  public Set<Characteristics> characteristics() {
    return Set.of(
      //Characteristics.CONCURRENT, we use regular lists, not its concurrent versions
      //Characteristics.UNORDERED, we do save the order
      Characteristics.IDENTITY_FINISH //the finisher function is the identity function and can be elided
    );
  }
}
