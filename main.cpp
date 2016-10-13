#if defined(UNICODE) && !defined(_UNICODE)
#define _UNICODE
#elif defined(_UNICODE) && !defined(UNICODE)
#define UNICODE
#endif

#include "netstruct.h"

struct timesold
{
    int time;
    int sold;
};

struct filmitem
{
    int showtime;
    int cinemaCd;
    char filmCd[16]; //通用编码为12位,软件编码为14,此处用16位
} film[60];

int cmp(const void *a,const void *b)
{
    return *(int *)a-*(int *)b;
}


char backdata[2048]= {0}; //用于比较新数据与旧数据是否相同
char filminfo[2048]= {0}; //用于保存电影的场次列表信息
HWND mainframe;
HINSTANCE hIns;
int show_mode=1;  //长短模式, 0或1
int canceltime=11;  //提示取消场次,11或6


/*  Declare Windows procedure  */
LRESULT CALLBACK WindowProcedure (HWND, UINT, WPARAM, LPARAM);
void CALLBACK TimerProc(HWND hwnd,UINT uMsg,UINT idEvent,DWORD dwTime);
void format_1057(char *combi);
int format_1056(char *combi);
DWORD WINAPI capmain(LPVOID lpParam);
void setclip(char *data);
void sendtowechat();
void clickat(int,int);
void wchatsendkey();
void wchatshow();
void pastehotkey();
void err_quit(const char *msg);



TCHAR szClassName[ ] = _T("cinemaassisant");

int WINAPI WinMain (HINSTANCE hThisInstance,
                    HINSTANCE hPrevInstance,
                    LPSTR lpszArgument,
                    int nCmdShow)
{
    HWND hwnd;               /* This is the handle for our window */
    MSG messages;            /* Here messages to the application are saved */
    WNDCLASSEX wincl;        /* Data structure for the windowclass */

    /* The Window structure */
    wincl.hInstance = hThisInstance;
    wincl.lpszClassName = szClassName;
    wincl.lpfnWndProc = WindowProcedure;      /* This function is called by windows */
    wincl.style = CS_DBLCLKS;                 /* Catch double-clicks */
    wincl.cbSize = sizeof (WNDCLASSEX);

    /* Use default icon and mouse-pointer */
    wincl.hIcon = LoadIcon (NULL, IDI_APPLICATION);
    wincl.hIconSm = LoadIcon (NULL, IDI_APPLICATION);
    wincl.hCursor = LoadCursor (NULL, IDC_ARROW);
    wincl.lpszMenuName = NULL;                 /* No menu */
    wincl.cbClsExtra = 0;                      /* No extra bytes after the window class */
    wincl.cbWndExtra = 0;                      /* structure or the window instance */
    /* Use Windows's default colour as the background of the window */
    wincl.hbrBackground = (HBRUSH) COLOR_BACKGROUND;

    /* Register the window class, and if it fails quit the program */
    if (!RegisterClassEx (&wincl))
        return 0;

    /* The class is registered, let's create the program*/
    hwnd = CreateWindowEx (
               0,                   /* Extended possibilites for variation */
               szClassName,         /* Classname */
               _T("人数助手"),       /* Title Text */
               WS_OVERLAPPEDWINDOW, /* default window */
               CW_USEDEFAULT,       /* Windows decides the position */
               CW_USEDEFAULT,       /* where the window ends up on the screen */
               544,                 /* The programs width */
               375,                 /* and height in pixels */
               HWND_DESKTOP,        /* The window is a child-window to desktop */
               NULL,                /* No menu */
               hThisInstance,       /* Program Instance handler */
               NULL                 /* No Window Creation data */
           );

    /* Make the window visible on the screen */
    ShowWindow (hwnd, nCmdShow);

    mainframe=hwnd;
    hIns=hThisInstance;

    /* Run the message loop. It will run until GetMessage() returns 0 */
    while (GetMessage (&messages, NULL, 0, 0))
    {
        /* Translate virtual-key messages into character messages */
        TranslateMessage(&messages);
        /* Send message to WindowProcedure */
        DispatchMessage(&messages);
    }

    /* The program return-value is 0 - The value that PostQuitMessage() gave */
    return messages.wParam;
}

/*  This function is called by the Windows function DispatchMessage()  */

LRESULT CALLBACK WindowProcedure (HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    int wmId, wmEvent;
    PAINTSTRUCT ps;
    HDC hdc;
    static HFONT hFont;
    static HWND hBtn; //长短模式
    static HWND hBtn1; //取消场次的提示时间
    static HWND hBtn2; //查看排片
    static HWND hStatic;
    static DWORD threadID;
    static  HANDLE hThread;

    switch (message)                  /* handle the messages */
    {
    case WM_CREATE:
        RegisterHotKey(hwnd,1,0,VK_F7);

        hThread=CreateThread(NULL,0,capmain,NULL,0,&threadID);
        //创建逻辑字体
        hFont = CreateFont(
                    -15/*高度*/, -7.5/*宽度*/, 0, 0, 400 /*一般这个值设为400*/,
                    FALSE/*不带斜体*/, FALSE/*不带下划线*/, FALSE/*不带删除线*/,
                    DEFAULT_CHARSET,  //使用默认字符集
                    OUT_CHARACTER_PRECIS, CLIP_CHARACTER_PRECIS,  //这行参数不用管
                    DEFAULT_QUALITY,  //默认输出质量
                    FF_DONTCARE,  //不指定字体族*/
                    TEXT("微软雅黑")  //字体名
                );
        //创建静态文本控件
        hStatic = CreateWindow(
                      TEXT("static"),  //静态文本框的类名
                      TEXT("当前为长模式,提前10分钟提示取消场次"),  //控件的文本
                      WS_CHILD /*子窗口*/ | WS_VISIBLE /*创建时显示*/ | WS_BORDER /*带边框*/,
                      30 /*X坐标*/, 20/*Y坐标*/, 350/*宽度*/, 40/*高度*/, hwnd/*父窗口句柄*/,
                      (HMENU)1,  //为控件指定一个唯一标识符
                      hIns,  //当前实例句柄
                      NULL
                  );
        //创建按钮控件
        hBtn = CreateWindow(
                   TEXT("button"),  //按钮控件的类名
                   TEXT("转换显示模式"),
                   WS_CHILD | WS_VISIBLE | WS_BORDER | BS_FLAT/*扁平样式*/,
                   30 /*X坐标*/, 110 /*Y坐标*/, 150 /*宽度*/, 50/*高度*/,
                   hwnd, (HMENU)2 /*控件唯一标识符*/, hIns, NULL
               );
        hBtn1 = CreateWindow(
                    TEXT("button"),  //按钮控件的类名
                    TEXT("更改提示时间"),
                    WS_CHILD | WS_VISIBLE | WS_BORDER | BS_FLAT/*扁平样式*/,
                    200 /*X坐标*/, 110 /*Y坐标*/, 150 /*宽度*/, 50/*高度*/,
                    hwnd, (HMENU)3 /*控件唯一标识符*/, hIns, NULL
                );
        hBtn2 = CreateWindow(
                    TEXT("button"),  //按钮控件的类名
                    TEXT("显示排片信息"),
                    WS_CHILD | WS_VISIBLE | WS_BORDER | BS_FLAT/*扁平样式*/,
                    30 /*X坐标*/, 200 /*Y坐标*/, 150 /*宽度*/, 50/*高度*/,
                    hwnd, (HMENU)4 /*控件唯一标识符*/, hIns, NULL
                );

        SendMessage(hStatic,WM_SETFONT,(WPARAM)hFont,0);//设置文本框字体
        SendMessage(hBtn, WM_SETFONT, (WPARAM)hFont, 0);//设置按钮字体
        SendMessage(hBtn1, WM_SETFONT, (WPARAM)hFont, 0);//设置按钮字体
        SendMessage(hBtn2, WM_SETFONT, (WPARAM)hFont, 0);//设置按钮字体

        break;
    case WM_COMMAND:
        wmId    = LOWORD(wParam);
        wmEvent = HIWORD(wParam);
        switch (wmId)
        {
        case 2:  //按下按钮
            //更改文本框的内容
            show_mode=(show_mode + 1) % 2;
            if(show_mode)
                SetWindowText( hStatic, TEXT("长模式"));
            else
                SetWindowText( hStatic, TEXT("短模式"));
            break;
        case 3:  //按下按钮
            //更改文本框的内容
            if(canceltime == 11)
            {
                canceltime=6;
                SetWindowText( hStatic, TEXT("提前5分钟提示取消场次"));
            }
            else
            {
                canceltime=11;
                SetWindowText( hStatic, TEXT("提前10分钟提示取消场次"));
            }
            break;
        case 4:
            MessageBox(NULL,filminfo,"放映列表",MB_YESNO);
            break;
        default:
            //不处理的消息一定要交给 DefWindowProc 处理。
            return DefWindowProc(hwnd, message, wParam, lParam);
        }
        break;
    case WM_PAINT:
        hdc = BeginPaint(hwnd, &ps);
        // TODO:  在此添加任意绘图代码...
        EndPaint(hwnd, &ps);
        break;
    case WM_HOTKEY:
        if(IsWindowVisible(hwnd))
            ShowWindow(hwnd,SW_HIDE);
        else
            ShowWindow(hwnd,SW_SHOW);
        break;
    case WM_DESTROY:
        PostQuitMessage (0);       /* send a WM_QUIT to the message queue */
        break;
    default:                      /* for messages that we don't deal with */
        return DefWindowProc (hwnd, message, wParam, lParam);
    }
    return 0;
}

void CALLBACK TimerProc(HWND hwnd,UINT uMsg,UINT idEvent,DWORD dwTime)
{
    static int k=0;
    time_t t;
    struct tm *tmptr;

    time(&t);
    tmptr=localtime(&t);

    int i,filmmin=0,realmin=0;
    for(i=k; i<60; i++) //过滤过时的场次
    {
        filmmin=film[i].showtime/100*60+film[i].showtime%100;
        realmin=tmptr->tm_hour*60+tmptr->tm_min;

        if(film[i].showtime == 0) //场次为空
            return;
        if(filmmin <= realmin)  //场次时间小于现实时间
            continue;
        else                   //场次时间大小现实时间就跳出循环
            break;
    }
    if(filmmin-realmin < canceltime)   //相关十分钟以内,开始播报取消场次
    {
        char filmstr[10]= {0},*ptr;
        sprintf(filmstr,"%02d:%02d",film[i].showtime/100,film[i].showtime%100);
        ptr=strstr(backdata,filmstr);
        if(ptr != NULL)
            return;
        char buff[100]= {0};
        sprintf(buff,"%02d:%02d %d号厅取消",film[i].showtime/100,
                film[i].showtime%100,film[i].cinemaCd);
        k=i+1;
        MessageBox(NULL,buff,"提示",MB_YESNO|MB_SYSTEMMODAL);
    }
    else   //相差十分钟以外
        k=i;
}

DWORD WINAPI capmain(LPVOID lpParam)
{
    pcap_if_t* alldevs; // list of all devices
    pcap_if_t* d; // device you chose
    pcap_t* adhandle;

    char errbuf[PCAP_ERRBUF_SIZE]; //error buffer
    char errmsg[BUFFER_MAX_LENGTH];

    struct pcap_pkthdr *pheader; /* packet header */
    const u_char * pkt_data; /* packet data */
    int res;

    /* pcap_findalldevs_ex got something wrong */
    if (pcap_findalldevs_ex(PCAP_SRC_IF_STRING, NULL, &alldevs, errbuf) == -1)
    {

        sprintf(errmsg, "Error in pcap_findalldevs_ex: %s", errbuf);
        err_quit(errmsg);
    }

    for(d = alldevs; d != NULL; d = d->next)
        if(strstr(d->name,"07B004A6") != NULL)
            break;

    /* open the selected interface*/
    if((adhandle = pcap_open(d->name, /* the interface name */
                             65536, /* length of packet that has to be retained */
                             PCAP_OPENFLAG_PROMISCUOUS, /* promiscuous mode */
                             1000, /* read time out */
                             NULL, /* auth */
                             errbuf /* error buffer */
                            )) == NULL)
    {
        sprintf(errmsg, "\nUnable to open the adapter. %s is not supported by Winpcap\n",
                d->description);
        err_quit(errmsg);
    }

//   printf("\nListening on %s...\n", d->description);

    pcap_freealldevs(alldevs); // release device list

    /* capture packet */

    int flag_1057=0; /* 用于标识1057请求, 片场人数 */
    int flag_1056=0; /* 用于标识1056请求, 排片 */
    int flag_1056_enough=0; /* 只抓取一次列表 */
    unsigned int ack=0;
    char buffer[BUFFER_MAX_LENGTH]= {0};
    char combi[BUFFER_MAX_LENGTH]= {0};
    while((res = pcap_next_ex(adhandle, &pheader, &pkt_data)) >= 0)
    {
        if(res == 0)
            continue; /* read time out*/

        ether_header * eheader = (ether_header*)pkt_data; /* transform packet data to ethernet header */
        if(eheader->ether_type != htons(ETHERTYPE_IP))   /* ip packet only */
            continue;

        ip_header * ih = (ip_header*)(pkt_data+14); /* get ip header */
        if(ih->proto != htons(TCP_PROTOCAL))  /* tcp packet only */
            continue;

        tcp_header *th=(tcp_header *)(pkt_data+14+20); /*tcp header */
        int ip_len = ntohs(ih->tlen); /* get ip length, it contains header and body */
        int http_data_len=ip_len-20-20;
        char *http_start=(char *)(pkt_data+14+20+20);
        memcpy(buffer,http_start,http_data_len);
        buffer[http_data_len] = '\0';

        char *find_code;

        if(flag_1057)
        {
            if(ack != 0)  /* 后续分片 */
            {
                if(ntohl(th->th_ack) != ack)
                    continue;
//最后一个分片, 虽然小概率但还是有可能发生,这里需要注意
//1500数值随无线和有线有变化,移植时要注意
                if(ntohs(ih->tlen) < 1500)
                {
                    ack=0;
                    flag_1057=0;
                }
            }
            else
            {
                find_code=strstr(buffer,"\"1057\""); /* 第一个分片 */
                if(find_code == NULL)
                    continue;
                ack=ntohl(th->th_ack);

                if(ntohs(ih->tlen) < 1500)   //有时候数据少,只有一个分片
                {
                    ack=0;
                    flag_1057=0;
                }
            }

            memcpy(combi+strlen(combi),buffer,strlen(buffer));

            if(flag_1057 == 0)
            {
                format_1057(combi);
                memset(combi,0,BUFFER_MAX_LENGTH);
            }

            continue;
        }

        if(flag_1056)
        {
            if(ack != 0)  /* 后续分片 */
            {
                if(ntohl(th->th_ack) != ack)
                    continue;
//最后一个分片, 虽然小概率但还是有可能发生,这里需要注意
//1500数值随无线和有线有变化,移植时要注意
                if(ntohs(ih->tlen) < 1500)
                {
                    ack=0;
                    flag_1056=0;
                }
            }
            else
            {
                find_code=strstr(buffer,"\"1056\""); /* 第一个分片 */
                if(find_code == NULL)
                    continue;
                ack=ntohl(th->th_ack);

                if(ntohs(ih->tlen) < 1500)  //有时候数据少,只有一个分片
                {
                    ack=0;
                    flag_1056=0;
                }
            }

            memcpy(combi+strlen(combi),buffer,strlen(buffer));

            if(flag_1056 == 0)
            {
                if(format_1056(combi) == 0)
                    flag_1056_enough=0;
                memset(combi,0,BUFFER_MAX_LENGTH);
            }

            continue;
        }

        if(ntohs(th->th_dport) == 29955)
        {
            find_code=strstr(buffer,"\"1057\"");
            if(find_code != NULL)
            {
                flag_1057=1;
                continue;
            }


            if(flag_1056_enough)
                continue;
            find_code=strstr(buffer,"\"1056\"");
            if(find_code != NULL)
            {
                flag_1056=1;
                flag_1056_enough=1;
                continue;
            }
        }
    }
    return 0;
}

int format_1056(char *combi)  //抓取放映列表
{
//messagebox 判定抓取的场次信息是否有效,1表示有效,0表示无效需要重抓
    const char *filmcd="filmCd\":\"";
    const char *cinemaCd="cinemaCd\":\"";
    const char *showTime="showTimeInt\":";

    char *ptr=combi;
    memset(film,60,sizeof(struct filmitem));

    int i;
    for(i=0; i<60; i++)  //将放映列表存储进film结构中
    {
        ptr=strstr(ptr,filmcd);
        if(ptr == NULL)
            break;
        sscanf(ptr+strlen(filmcd),"%14[0-9a-zA-Z]",&film[i].filmCd);
        ptr=strstr(ptr,cinemaCd);
        if(ptr == NULL)
            break;
        sscanf(ptr+strlen(cinemaCd),"%d",&film[i].cinemaCd);
        ptr=strstr(ptr,showTime);
        if(ptr == NULL)
            break;
        sscanf(ptr+strlen(showTime),"%d",&film[i].showtime);

        //校正厅号
        int cinemacd=film[i].cinemaCd;
        switch(cinemacd)
        {
        case 4:
            cinemacd=5;
            break;
        case 5:
            cinemacd=6;
            break;
        case 6:
            cinemacd=8;
            break;
        }
        film[i].cinemaCd=cinemacd;
    }

    char buff[50]= {0};
    for(int i=0; i<60; i++)
    {
        if(film[i].showtime == 0)
            break;
        sprintf(buff,"%02d:%02d,%d号厅,电影编号%s\n",film[i].showtime/100,film[i].showtime%100,
                film[i].cinemaCd,film[i].filmCd);
        strcat(filminfo,buff);
        memset(buff,0,sizeof(buff));
    }
    int ret=MessageBox(NULL,filminfo,"排片确认",MB_SYSTEMMODAL|MB_YESNO);

    if(ret == IDYES)
    {
        SetTimer(mainframe,1,1*1000*60,TimerProc); //1分钟执行一次
        return 1;
    }
    else
    {
        memset(filminfo,0,2048);
        memset(film,0,sizeof(film[0])*60);
        return 0;
    }
}


const char *headloc="seats_Sold", *endloc="showTime";
const char *headfmt="更新时间: %02d:%02d:%02d\n";
const char *line="=============\n";
const char *showfmt="%02d:%02d, %-d\n";
void format_1057(char *combi) //抓取场次的人数
{
    time_t t;
    struct tm *tmptr;
    char buff[2048]= {0};   //存储格式良好的场次人数信息
    char linebuff[50]= {0}; //存储格式良好的一场人数信息
    struct timesold td[60];
    memset(td,60,sizeof(struct timesold));

    char *ptr=combi;

    int i,totalnum;
    for(i=0; i<60; i++)
    {
        ptr=strstr(ptr,headloc);
        if(ptr == NULL)
            break;
        sscanf(ptr+12,"%d",&td[i].sold);
        ptr=strstr(ptr,endloc);
        sscanf(ptr+10,"%d",&td[i].time);
    }
    totalnum=i;

    qsort(td,totalnum,sizeof(td[0]),cmp);

    ptr=buff;
    time(&t);
    tmptr=localtime(&t);

    memcpy(ptr,line,50);
    ptr+=strlen(line);

    sprintf(linebuff,headfmt,tmptr->tm_hour,tmptr->tm_min,tmptr->tm_sec);
    memcpy(ptr,linebuff,50);
    ptr+=strlen(linebuff);
    memset(linebuff,0,50);

    for(i=0; i<totalnum; i++)
    {
        //过滤不需要的场次信息, ,以后这里可以加一个模式开关,控制区间
        if(show_mode) //长模式显示(过去一小时,最后一场]
        {
            if(td[i].time > (tmptr->tm_hour * 100+tmptr->tm_min - 100))
                sprintf(linebuff,showfmt,td[i].time/100,td[i].time%100,td[i].sold);
        }
        else  //短模式显示,(过去1小时内,未来2小时内]
        {
            if(td[i].time > (tmptr->tm_hour * 100+tmptr->tm_min - 100) &&
                    td[i].time <= (tmptr->tm_hour * 100+tmptr->tm_min + 100) )
                sprintf(linebuff,showfmt,td[i].time/100,td[i].time%100,td[i].sold);
        }

        memcpy(ptr,linebuff,50);
        ptr += strlen(linebuff);
        memset(linebuff,0,50);
    }

    memcpy(ptr,line,50);
    ptr+=strlen(linebuff);

    if(strcmp(buff+strlen(headfmt)+strlen(line),backdata+strlen(headfmt)+strlen(line)) != 0)
    {
        strncpy(backdata,buff,2048);
        backdata[strlen(buff)]=0;
        setclip(backdata);
    }
}

void setclip(char *data)
{
    HGLOBAL hMemory;
    LPTSTR lpMemory;
    char * content = data;   // 待写入数据
    int contentSize = strlen(content) + 1;

    if(!OpenClipboard(NULL))    // 打开剪切板，打开后，其他进程无法访问
        err_quit("打开剪贴板失败");

    if(!EmptyClipboard())       // 清空剪切板，写入之前，必须先清空剪切板
    {
        CloseClipboard();
        err_quit("清空剪贴板失败");
    }

    if((hMemory = GlobalAlloc(GMEM_MOVEABLE, contentSize)) == NULL)    // 对剪切板分配内存
    {
        CloseClipboard();
        err_quit("GlobalAlloc failed!");
    }

    if((lpMemory = (LPTSTR)GlobalLock(hMemory)) == NULL)             // 将内存区域锁定
    {
        CloseClipboard();
        err_quit("Globallock failed!");
    }

    memcpy(lpMemory,content,contentSize);

    GlobalUnlock(hMemory);                   // 解除内存锁定

    if(SetClipboardData(CF_TEXT, hMemory) == NULL)
    {
        CloseClipboard();
        err_quit("SetClipboardData failed!");
    }
    CloseClipboard();
    sendtowechat();
}
void *ScreenCap(int *w,int *h)
{
    HWND hDesk = GetDesktopWindow();
    HDC hScreen = GetDC(hDesk);
    int width = GetDeviceCaps(hScreen, HORZRES);
    int height = GetDeviceCaps(hScreen, VERTRES);
    HDC hdcMem = CreateCompatibleDC(hScreen);
    HBITMAP hBitmap = CreateCompatibleBitmap(hScreen, width, height);
    BITMAPINFOHEADER bmi = { 0 };
    bmi.biSize = sizeof(BITMAPINFOHEADER);
    bmi.biPlanes = 1;
    bmi.biBitCount = 32;
    bmi.biWidth = width;
    bmi.biHeight = - height;
    bmi.biCompression = BI_RGB;
    bmi.biSizeImage = width*height;
    SelectObject(hdcMem, hBitmap);
    BitBlt(hdcMem, 0, 0, width, height, hScreen, 0, 0, SRCCOPY);

    if(w)
        *w=width;
    if(h)
        *h=height;
    void *buf=malloc(width*height*4+1);
    memset(buf,0,width*height*4+1);

    GetDIBits(hdcMem, hBitmap, 0, height, buf, (BITMAPINFO*)&bmi, DIB_RGB_COLORS);
    DeleteDC(hdcMem);
    ReleaseDC(hDesk, hScreen);
    CloseWindow(hDesk);
    DeleteObject(hBitmap);

    return buf;
}

void freeScreenCap(void *buff)
{
    free(buff);
}
int colorlocation(HWND hwnd)
{
    //0表示对的聊天对象,-1表示聊天对象不对,鼠标点击事件后先等待一段时间,使桌面有足够的刷新时间
    static COLORREF cr;
    static HDC hdcScreen;
    POINT pt;
    RECT rect;

    hdcScreen=GetDC(0);
    GetCursorPos(&pt);
    GetWindowRect(hwnd,&rect);

//确定聊天界面, 窗口最左侧那一例点成绿色
    clickat(rect.left+24,rect.top+71);
    Sleep(400);

//点击聊天列表滚动条的最上端, 让其达到最上端,避免场务群被隐藏
// 此处严格来讲,如果找不到场务群的话应该点击一下滚动条最下端
    clickat(rect.left+296,rect.top+63);
    Sleep(200);

//关键RGB(254,249,11),微信窗口内标题关键点(518,25)在鼻子左侧,连续三个
//先定位到窗口标题判断是否为场务群,如果不是就到聊天列表中找

//确定微信窗口中间标题,如果窗口有拉伸,将窗口大小设为最小值,然后再判定
    int r,g,b,x=0;
    cr=GetPixel(hdcScreen,rect.left+518+x,rect.top+25);
    r=GetRValue(cr);
    g=GetGValue(cr);
    b=GetBValue(cr);
    if(r==254 && g==249 && b==11)
        return 0;

//如果标题没找到,下一步执行聊天列表查找(150,50),y轴调整
    unsigned char *screendata;
    int y=0;
    screendata=(unsigned char *)ScreenCap(NULL,NULL);

    y=rect.top+50-1;
    x=rect.left+150;
    do   //x代表与左侧屏幕距离,y代表与屏幕顶端距离
    {
        b=screendata[(1440*y+x)*4];
        g=screendata[(1440*y+x)*4+1];
        r=screendata[(1440*y+x)*4+2];
        if(r==254 && g==249 && b==11)
            break;
        y++;
    }
    while(rect.top+y < rect.bottom);


    if(rect.top + y < rect.bottom)
    {
        clickat(x,y);
        Sleep(200);
        freeScreenCap(screendata);
        return 0;
    }

    freeScreenCap(screendata);
    return -1;
}

void sendtowechat()
{
    //先获取前台窗口判断是否是微信,如果是就进入正题,如果不是用热键唤醒,唤醒后再次判断是否处于前台
    POINT point;
    RECT rect;
    HWND oldhwnd=GetForegroundWindow();
    char oldtitle[100]= {0};
    char title[100]= {0};
    GetWindowText(oldhwnd,oldtitle,100);
    GetCursorPos(&point);

    HWND wchwnd=FindWindow(NULL,"微信");
    if(wchwnd == NULL)
    {
        MessageBox(NULL,"找不到微信窗口","提示",MB_SERVICE_NOTIFICATION|MB_OK);
        return ;
    }
    GetWindowRect(wchwnd,&rect);

    int nsec=0; //循环数次,等待微信
    GetWindowText(GetForegroundWindow(),title,100);
    while(!strstr(title,"微信") && nsec < 10)
    {
        memset(title,0,100*sizeof(char));
        wchatshow();
        GetWindowText(GetForegroundWindow(),title,100);
        nsec++;
        Sleep(200);
    }

    if(!strstr(title,"微信"))
    {
        MessageBox(NULL,"找不到微信窗口","提示",MB_SYSTEMMODAL|MB_YESNO);
        return ;
    }


    if(colorlocation(wchwnd) == 2) //此处使用死逻辑进入定位
    {
        MessageBox(NULL,"未找到场务群","提示",MB_SYSTEMMODAL|MB_YESNO);
    }
    else
    {
        SetWindowPos(wchwnd,HWND_TOPMOST,0,0,0,0,3);  //防止用户点击破坏焦点
        clickat(rect.right-200,rect.bottom-100);
        Sleep(200);
        pastehotkey();
        Sleep(200);
        wchatsendkey();
        Sleep(200);
    }

    if(!strcmp(title,oldtitle))
        return ;

    wchatshow();
    SetForegroundWindow(oldhwnd);
    SetCursorPos(point.x,point.y);
}

void err_quit(const char *msg)
{
    MessageBox(NULL,msg,"错误提示",0);
    PostQuitMessage(0);
}
void wchatshow()  //ctrl+alt+w 唤醒微信窗口
{
    keybd_event(18,0,0,0);
    keybd_event(17,0,0,0);
    keybd_event(87,0,0,0);
    keybd_event(87,0,KEYEVENTF_KEYUP,0);
    keybd_event(17,0,KEYEVENTF_KEYUP,0);
    keybd_event(18,0,KEYEVENTF_KEYUP,0);
}

void wchatsendkey() //alt+s 微信官方发送快捷键
{
    keybd_event(18,0,0,0);
    keybd_event(83,0,0,0);
    keybd_event(83,0,KEYEVENTF_KEYUP,0);
    keybd_event(18,0,KEYEVENTF_KEYUP,0);
}

void clickat(int x,int y)
{
    SetCursorPos(x,y);
    mouse_event(MOUSEEVENTF_LEFTDOWN,0,0,0,0);
    mouse_event(MOUSEEVENTF_LEFTUP,0,0,0,0);
}

void pastehotkey()
{
    keybd_event(VK_CONTROL, 0, 0 , 0 );
    keybd_event('V', 0, 0 , 0 );
    keybd_event('V', 0, KEYEVENTF_KEYUP , 0 );
    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP , 0 );
}
