from kaggle_environments import make
env = make("rps", configuration={"episodeSteps": 1000}, debug = True)

env.reset()
env.run(["beat_a_human.py", "beat_a_human.py"])
env.render(mode="ipython", width=550, height=500)