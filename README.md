# ppt-render-html
## Why  
I needed a simple local home web page which could be edited in a graphical WYSIWYG way, for a kid-friendly dashboard.    
  
Here's one way: create a PowerPoint presentation, embed it into a simple html page, and voilà!  
  
Several problems, though.  
	1) Microsoft did not care about high-resolution screens. So what looks like a crisp image in PowerPoint turns into a blurred embarassment in an embedding.  
	2) The resulting page is not adaptive at all. The dimensions are fixed.  
  3) Microsoft can change the link to the shared presentation, or the embedding stops working altogether.    
All in all, not the greatest feeling when interacting with the page.  
<p align="center">
  <img src="docs/page1.png" width="1473">
</p>  
<p align="center">
  <img src="docs/refuse.png" width="200">
</p>  

I wanted something like ppt exported to html, but this functionality has been long gone from PowerPoint.  
  
An appealing alternative would be using Miro. But their link buttons look cluttered to me and they can't be disabled.   
A long time ago, in a country far, far away, in a post not saved by Web Archive, I read a nice write-up by Miro that described how they migrated rendering of their content to HTML canvas.  
So I decided to make my own export of slides to HTML canvas.  
  
My main requirements:  
	1) The page is 100% identical visually to the original slide  
	2) Links are clickable  
	3) The page scales to the window  
  4) The page is rendered taking the high resolution screen into account   
  <p align="center">
  <img src="docs/page2.png" width="1473">
</p>  
