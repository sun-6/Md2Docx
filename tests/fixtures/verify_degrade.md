# verify 退化场景测试

正常公式（应保留为原生公式）：

$$\int_0^1 x^2 \, dx = \frac{1}{3}$$

无法解析的宏（会退化为原始 TeX 文本——verify 应报告公式数不足）：

$$\thismacrodoesnotexist{a}{b} + \alsobogus{x}$$
