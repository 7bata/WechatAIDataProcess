from docx import Document

# Load word
for i in range(100):
    document = Document("D:/WechatAI/WebContent/Word/file_{}.docx".format(i+1))
    # Get the text from doc
    document_text = []
    for paragraph in document.paragraphs:
        document_text.append(paragraph.text)
    print("Processing file_{}".format(i+1))
    # Save text doc
    with open("D:/WechatAI/WebContent/TXT/file_{}.txt".format(i+1), "w", encoding="utf-8") as file:
        file.write("\n".join(document_text))

print("Finished")
