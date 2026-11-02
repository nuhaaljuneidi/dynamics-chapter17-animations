# Dynamics Chapter 17 Animations

Animated GIFs from the Chapter 17 (Planar Kinematics of Rigid Bodies) lecture slides, one section per animation.

Site: https://nuhaaljuneidi.github.io/dynamics-chapter17-animations/

Each animation has its own link, for example `#problem12` or `#ic-cases-rolling`.

## Updating

After adding or changing GIFs in the PowerPoint deck, save it and run:

```
python update.py
```

or point it at a different deck:

```
python update.py path/to/deck.pptx
```

The script reads every GIF in the deck, names it from its slide title, rewrites `gifs/` and `index.html`, and prints the list. Then commit and push:

```
git add -A
git commit -m "Update animations"
git push
```

To choose a name yourself, add the slide title and the picture's name in PowerPoint to `NAMES` at the top of `update.py`.
