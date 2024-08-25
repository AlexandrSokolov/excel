### extend excel writer with new java types

extend currently supported java types for cell values.

Currently it is only:
- Double -> CellType.NUMERIC
- BigDecimal -> CellType.NUMERIC
- String -> CellType.STRING
- null -> CellType.BLANK

Add support:
- Date time java types -> ?

### extend excel writer with `attribute -> value` mapping

now excel writer accepts rows as a mappings between column letter and its value

Allow to accept attribute -> value mapping with the same configuration logic applied, as for Excel reader

