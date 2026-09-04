"""物流跟踪自动更新：读物流跟踪表格里"未完成"的运单，按物流商分组、并行去各个货代平台查
最后路由，更新回表格的"最后流水"/"是否有更新"两列。业务逻辑见 tracking_pipeline.py，界面见
panel.py，各货代平台的查询实现在 platforms/ 下（照搬自独立脚本 logistics_tracking/，见
platforms/registry.py 顶部说明）。
"""
