#ifndef WZWORD_H
#define WZWORD_H

#include <QAxObject>
#include <QString>

/*
 *  类：WzWord
 *  用途：Qt操作Word文档
 *  作者：欧阳伟
 *  日期：2018-8-1
 *  示例：
 *      ①在*.pro文件中添加：QT += axcontainer
 *      ②在word文件中添加标签
 *      ③示例代码：
 *          A:(简单场景，无需复用，无需close，析构时会自动关闭并释放内存)
 *          WzWord w("D:/template.docx");
 *          if(w.open())
 *          {
 *              w.insertTextIntoLabel("Header","静夜思");
 *              w.insertTextIntoLabel("Paragraph1","窗前明月光");
 *              w.insertTextIntoLabel("Paragraph2","疑是地上霜");
 *              w.insertTextIntoLabel("Paragraph3","举头望明月");
 *              w.insertTextIntoLabel("Paragraph4","低头思故乡");
 *              w.insertPictureIntoLabel("picture1","D:/project/图片素材/man.png");
 *              w.insertPictureIntoLabel("picture2","D:/project/图片素材/women.png",300,200);
 *              //w.save();
 *              w.saveAs("D:/LiBai.docx");
 *          }
 *
 *          B:（相对复杂场景，一个对象，多处使用，需close）
 *          WzWord w;
 *          w.setFileName("D:/template.docx");
 *          if(w.open())
 *          {
 *              w.insertTextIntoLabel("Header","静夜思");
 *              w.insertTextIntoLabel("Paragraph1","床前明月光");
 *              w.insertTextIntoLabel("Paragraph2","疑是地上霜");
 *              w.insertTextIntoLabel("Paragraph3","举头望明月");
 *              w.insertTextIntoLabel("Paragraph4","低头思故乡");
 *              w.insertPictureIntoLabel("Picture1","D:/a.jpg");
 *              w.insertPictureIntoLabel("Picture2","D:/b.png",200,200);
 *              w.saveAs("D:/静夜思.docx");
 *              w.close();
 *          }
 *
 *  说明：①传入绝对路径。②请用*.docx格式(我的测试格式，其他格式不保证稳定)
 */

class WzWord
{
public:
    WzWord();
    WzWord(const QString fileName);
    ~WzWord();

    //打开，不存在则创建，默认窗体不可见，不显示警告，如无特殊需要，请保持默认
    bool open(bool visible=false,bool displayAlerts=false);
    //设置visible，true: 可视，false: 隐藏
    bool setVisible(bool visible);
    //保存
    bool save();
    //另存为
    bool saveAs(const QString fileName);
    //标签处插入内容
    bool insertTextIntoLabel(const QString label,const QString content);
    //标签处插入图片
    bool insertPictureIntoLabel(const QString label,const QString fileName);
    //标签处插入图片,自定义图片宽度和高度
    bool insertPictureIntoLabel(const QString label,const QString fileName,const QVariant width,const QVariant height);
    //关闭、释放资源
    void close();
    //设置文件名，搭配无参构造函数使用
    void setFileName(const QString fileName);

private:
    //释放资源函数
    void release();

private:
    bool isOpened;
    QString fileName;//文件名

    QAxObject *word;//word对象
    QAxObject *documents;//所有文档对象
    QAxObject *currentDocument;//当前文档对象
};

#endif // WZWORD_H
