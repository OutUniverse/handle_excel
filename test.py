import asyncio

async def test(i):
    print(i)

async def main():
    for i in range(1, 10):
        await test(i)

asyncio.run(main())


if __name__ == "__main__":
    print("yes")