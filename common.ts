export const sleep = (time: number) =>
  new Promise((res) => {
    console.info(`wait for ${time} seconds`);
    setTimeout(() => {
      console.info(`complete wait for ${time} seconds`);
      res(time);
    }, time);
  });
