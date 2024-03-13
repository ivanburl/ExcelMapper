# ExcelMapper

Mapper works similar as Model Mapper, 
but for object to excel worksheet mapping. Solution is based on Annotations and Reflection. 

# TODO

- ## Features 
- [ ] Support of Map interface
- [x] Support of Collection interface
- [x] Support of custom styles
- [ ] Cycles detection 
  - [ ] Exception throwing
  - [ ] Resolution based on nulls as stop points
- [ ] Adjust column width for better preview
- [ ] Better naming conventions
- [ ] Providers of data in case of null
  - [ ] Automatically detect no args constructor
- [ ] Configuration using builder design patter

- ## Tests

- [x] Simple integration tests
- [ ] Unit tests 

- ## Implementation

- [x] Fast Excel [link](https://github.com/dhatim/fastexcel)
- [ ] Apache POI [link](https://github.com/apache/poi)

- ## Benchmarks

- [ ] Benchmark using JMH (mapper speed not mapper creation) 