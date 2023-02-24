def coinChange1(coins, amount):
    memory = {}

    def helper(n):
        if n in memory:
            return memory[n]
        if n == 0:
            return 0
        if n < 0:
            return -1

        result = float('inf')
        for coin in coins:
            tmp = helper(n - coin)
            if tmp < 0:
                continue
            result = min(result, tmp+1)
        memory[n] = result if result != float('inf') else -1
        return memory[n]

    return helper(amount)

def coinChange2(coins, amount):
    memory = [0] + [float('inf')]*(amount)
    for i in range(1, amount+1):
        for coin in coins:
            if i - coin < 0:
                continue
            memory[i] = min(memory[i - coin]+1, memory[i])

    return memory[amount] if memory[amount] != float('inf') else -1

if __name__ == "__main__":
    print(coinChange1([2,5],19))
    print(coinChange2([2,5],19))
