import multiprocessing
import os
import asyncio

def cpu_bound_task():
    print(f"子进程 {os.getpid()} 正在执行CPU密集型任务")
    sum = 0
    for i in range(1000000):
        sum += i
    print(f"子进程 {os.getpid()} CPU密集型任务完成，结果：{sum}")
    

def cpu2():
    print(a)

async def main():
    process = multiprocessing.Process(target=cpu_bound_task)
    process.start()
    process.join()  # 等待进程完成

if __name__ == '__main__':
    # global a
    a=5    
    loop=asyncio.get_event_loop()
    loop.run_until_complete(main())
    print("主进程继续执行")
    cpu_bound_task()
    
