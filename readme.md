# 保存原来的格式。

# 是否要计算段落，只需要往前面增加标识即可，不需要用标签标识。

# 表格的话，
每过一次table，清空一下。arr map 记录的对象，不然累加的话，索引就对不上了，
因为他不是excel


# 需要解决的是，run 分割了 占位符与 tag。导致无法寻找的问题
- 解决方案1：直接合成一个run。（之前已经完成，但是无法保留格式）
- 解决方案2: 合并寻找.....，不知道格式是否会丢失