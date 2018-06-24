package com.zsf.attendance.utils;

/**
 * Created by 张胜锋
 * On 2018/6/12 8:08
 * Description:
 */
public class CommUtil {

    //默认开始上班打卡时间为早上9:00  计算规则:  小时*3600 + 分*60   32400 = 9*3600+0*60
    private static long DEFAULTSTARTWORKTIME = 32400;
    //浮动打卡时间  9.30打卡也是可以的。但是下班打卡时间要顺延半个小时
    private static long FLOATWORKTIME = 34200;
    private static String NORECORD = "未打卡";


    //计算加班时长
    private static long getTime(String time) {
        //定义一个long类型的加班时长
        long worktime = 0;
        int length = time.length();
        switch (length){
            case 3 :   //未打卡   传入字符串:未打卡  暂时不计算加班时长
                if (NORECORD.equals(time)){
                    return worktime;
                }
                break;
            case 6 :  //一天只打卡一次，比如说:09:16-     或者：  -18:20

                break;
            case 11: //一天打卡两次，比如说：09:09-20:36

                break;
            case 17: //一天打卡三次，比如说:09:15-21:57-22:00  下班时间按最晚打卡时间计算

                break;
            case 23:  //一天打卡四次，比如说:  08:39-08:41-08:46-16:07  你没看错，这就是逗比人事妹纸的打卡记录，满满的都是坑爹。

                break;
            default:

                break;
        }





        return worktime;
    }

}