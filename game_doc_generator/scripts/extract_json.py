import json
from read_docx_enhanced import read_docx_enhanced

data = read_docx_enhanced(r"g:\zmd works\skills\S神石玩法.docx", 
                          extract_images_flag=False)

# 保存为JSON
with open(r"g:\zmd works\skills\temp_images\shenshi\document_data.json", 
          "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("✓ JSON文件已保存")
print(f"✓ 共{len(data['content'])}个内容项")

# 统计包含图片的段落
image_paras = [item for item in data['content'] if item.get('has_images')]
print(f"✓ 包含图片的段落: {len(image_paras)}个")
