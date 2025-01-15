import threading
import time

def cpu_bound_task(num):
    """一个计算密集型任务"""
    sum = 0
    for i in range(num):
        sum += i * i

def main():
    start_time = time.time()
    
    # 创建线程列表
    threads = []
    for _ in range(8):  # 创建8个线程
        thread = threading.Thread(target=cpu_bound_task, args=(50000000,))
        threads.append(thread)
        thread.start()
    
    # 等待所有线程完成
    for thread in threads:
        thread.join()
    
    end_time = time.time()
    print(f"多线程执行时间：{end_time - start_time}秒")

if __name__ == "__main__":
    main()