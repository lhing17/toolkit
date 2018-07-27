
##一、常用命令
ls
cd
cp
mv
clear
history
ps -ef | grep tomcat
ll -h
tail -500f XXX
tail -n 10000 XXX > check.out
vim XXX
mkdir XXX
du -sh
df -h
top
netstat -ano
grep -riwv 0000 .
whereis XXX
date


##二、安装软件相关
yum install
yum list
rpm -ivh
rpm -qa | grep XXX
rpm -e XXX

##三、变更权限相关
chown -R mysql:mysql mysql/
chmod +X jweb.jar
chmod 777 jweb.jar

##四、用户、用户组相关
groupadd
useradd
