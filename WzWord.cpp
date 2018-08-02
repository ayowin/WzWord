#include "WzWord.h"

#include <QDebug>
#include <windows.h>
#include <QFile>
#include <QDir>

WzWord::WzWord()
{
    qDebug()<<"WzWord: 构造函数WzWord()";

    OleInitialize(0);//当前线程初始化COM库并设置并发模式STA（single-thread apartment——单线程单元），务必保证第一条语句

    word = NULL;
    documents = NULL;
    currentDocument = NULL;

    isOpened = false;
}

WzWord::WzWord(const QString fileName)
{
    qDebug()<<"WzWord: 构造函数WzWord(const QString fileName)";

    OleInitialize(0);//当前线程初始化COM库并设置并发模式STA（single-thread apartment——单线程单元）

    word = NULL;
    documents = NULL;
    currentDocument = NULL;

    isOpened = false;
    this->fileName = fileName;
}

WzWord::~WzWord()
{
    qDebug()<< "WzWord: ~WzWord()";
    release();

    OleUninitialize(); //关闭当前线程的COM库并释放相关资源，务必保证最后一条语句
}

bool WzWord::open(bool visible, bool displayAlerts)
{
    qDebug()<<"WzWord: open(bool visible=false, bool displayAlerts=false)";

    if(fileName.isEmpty())
    {
        qDebug()<<"打开失败，文件名为空";
        return false;
    }

    QFile file(fileName);
    if (!file.exists())
    {
        //文件不存在则创建
        if(!file.open(QIODevice::WriteOnly))
        {
            qDebug()<<"WzWord: 文件创建失败";
            file.close();
            return false;
        }
        file.close();
    }
    else
    {
        //存在，但不能以写入状态打开
        if(!file.open(QIODevice::ReadWrite))
        {
            qDebug()<<"WzWord: 打开失败，文件正在被打开，请先关闭";
            file.close();
            return false;
        }
        file.close();
    }

    word = new QAxObject("Word.Application"); //word对象
    word->dynamicCall("SetVisible(bool)", visible); //false,不显示窗体
    word->setProperty("DisplayAlerts", displayAlerts); //false,不显示警告
    documents = word->querySubObject("Documents"); //所有文档对象
    currentDocument = documents->querySubObject("Open(const QString &)", fileName); //当前文档对象
    if(currentDocument == NULL)
    {
        qDebug()<<"WzWord: 打开失败，当前文档未找到";
        return false;
    }

    isOpened = true;
    return true;
}

bool WzWord::setVisible(bool visible)
{
    qDebug()<<"WzWord: setVisible(bool visible)";

    if(word == NULL)
    {
        qDebug()<<"WzWord: 设置visible失败，word对象为空";
        return false;
    }
    else
    {
        word->dynamicCall("SetVisible(bool)", visible);
        return true;
    }
}

bool WzWord::save()
{
    qDebug()<< "WzWord: save()";

    if(!isOpened)
    {
        qDebug()<<"WzWord: 保存失败，文件未打开";
        return false;
    }

    QFile file(fileName);
    if(file.exists())
    {
        //存在则直接保存
        currentDocument->dynamicCall("Save()");
    }
    else
    {
        //不存在则另存为
        return this->saveAs(fileName);
    }

    return true;
}

bool WzWord::saveAs(QString fileName)
{
    qDebug()<<"WzWord: saveAs(QString fileName)";

    if(!isOpened)
    {
        qDebug()<<"WzWord: 另存为失败，文件未打开";
        return false;
    }

    //如果文件已存在，并且正在打开中
    QFile file(fileName);
    if(file.exists())
    {
        if(!file.open(QIODevice::ReadWrite))
        {
            qDebug()<<"WzWord: 另存为失败，文件正在打开中，请先关闭文件";
            return false;
        }
        else
        {
            file.close();
        }
    }

    currentDocument->dynamicCall("SaveAs(const QString &)",QDir::toNativeSeparators(fileName));
    return true;
}

bool WzWord::insertTextIntoLabel(QString label, QString content)
{
    qDebug()<<"WzWord: insertTextIntoLabel(QString label, QString content)";

    if(!isOpened)
    {
        qDebug()<<"WzWord: 插入文字失败，文件未打开";
        return false;
    }

    //获取标签对象
    QAxObject *label1 = currentDocument->querySubObject("Bookmarks(QString)",label);

    if(label1 == NULL)
    {
        qDebug()<<"WzWord: 插入失败，标签对象为空，获取标签失败";
        return false;
    }

    label1->dynamicCall("Select(void)");
    label1->querySubObject("Range")->setProperty("Text",content);

    return true;
}

bool WzWord::insertPictureIntoLabel(const QString label, const QString fileName)
{
    qDebug()<<"WzWord: insertPictureIntoLabel(const QString label, const QString fileName)";

    if(!isOpened)
    {
        qDebug()<<"WzWord: 插入图片失败，文件未打开";
        return false;
    }

    //获取标签对象
    QAxObject *label1 = currentDocument->querySubObject("Bookmarks(QString)",label);

    if(label1 == NULL)
    {
        qDebug()<<"WzWord: 插入失败，标签对象为空，获取标签失败";
        return false;
    }

    //检查图片是否存在
    QFile file(fileName);
    if(!file.exists())
    {
        qDebug()<<"WzWord: 插入失败，图片不存在";
        return false;
    }

    label1->dynamicCall("Select(void)");
    QAxObject *picture = currentDocument->querySubObject("InlineShapes");
    picture->dynamicCall("AddPicture(const QString,QVariant,QVariant,QVariant)",
                         fileName,
                         false,
                         true,
                         label1->querySubObject("Range")->asVariant());
    delete picture;

    return true;
}

bool WzWord::insertPictureIntoLabel(const QString label, const QString fileName, const QVariant width, const QVariant height)
{
    qDebug()<< "WzWord: insertPictureIntoLabel(const QString label, const QString fileName, const QVariant width, const QVariant height)";

    if(!isOpened)
    {
        qDebug()<<"WzWord: 插入图片失败，文件未打开";
        return false;
    }

    //获取标签对象
    QAxObject *label1 = currentDocument->querySubObject("Bookmarks(QString)",label);

    if(label1 == NULL)
    {
        qDebug()<<"WzWord: 插入失败，标签对象为空，获取标签失败";
        return false;
    }

    //检查图片是否存在
    QFile file(fileName);
    if(!file.exists())
    {
        qDebug()<<"WzWord: 插入失败，图片不存在";
        return false;
    }

    label1->dynamicCall("Select(void)");
    QAxObject *inlineShapes = currentDocument->querySubObject("InlineShapes");
    QAxObject *shape = inlineShapes->querySubObject("AddPicture(const QString&,QVariant,QVariant,QVariant)",
                                                    fileName,
                                                    false,
                                                    true,
                                                    label1->querySubObject("Range")->asVariant());
    shape->setProperty("Width",width); //设置图片宽度
    shape->setProperty("Height",height); //设置图片高度

    delete shape;
    delete inlineShapes;

    return true;
}

void WzWord::close()
{
    qDebug()<< "WzWord: close()";

    release();
}

void WzWord::setFileName(const QString fileName)
{
    qDebug()<< "setFileName(const QString fileName)";
    this->fileName = fileName;
}

void WzWord::release()
{
    //关闭文件
    isOpened = false;
    //释放当前文档资源
    if(currentDocument != NULL)
    {
        currentDocument->dynamicCall("Close(QVariant)",0);
        delete currentDocument;
        currentDocument = NULL;
    }
    //释放所有文档资源
    if(documents != NULL)
    {
        documents->dynamicCall("Close(QVariant)",0);
        delete documents;
        documents = NULL;
    }
    //关闭word
    if(word != NULL)
    {
        word->dynamicCall("Quit(QVariant)",0);
        delete word;
        word = NULL;
    }
}
