<!DOCTYPE html>
<html <?php language_attributes(); ?>>
<head>
<meta charset="<?php bloginfo( 'charset' ); ?>" />
<meta name="viewport" content="width=device-width" />
<title><?php wp_title( '', true ); ?></title>

<link rel="profile" href="http://gmpg.org/xfn/11" />
<link rel="pingback" href="<?php bloginfo( 'pingback_url' ); ?>" />
<link rel="icon" href="/favicon/favicon.svg" type="image/svg+xml">
<link rel="icon" href="/favicon/favicon.png" type="image/png">
<link rel="apple-touch-icon" href="/favicon/favicon.png">



<?php wp_head(); ?>

</head>

<body <?php body_class(); ?>>


<?php
	/*-----------------------------------------------------------------------------------*/
	/* Start header
	/*-----------------------------------------------------------------------------------*/
$icon_arrow = get_stylesheet_directory() . '/img/arrow.svg';
$logo_svg = get_stylesheet_directory() . '/img/logo.svg';
$trud_svg = get_stylesheet_directory() . '/img/trudnetru.svg';
$order_svg = get_stylesheet_directory() . '/img/oreder.svg';

?>

<header id="masthead" class="site-header" role="banner">
	<div class="container">
	    <a href="/"><img src="<?php echo get_template_directory_uri() ?>/img/trudnetru.svg" alt="Охрана труда в Кирове"></a>
	    <hr>
		
		
		<div id="brand">
		        <? if ( is_single() ) :
        echo '<h3 class="site-title"> ';
    else :
        echo '<h1 class="site-title"> ';
    endif; ?>
		    
			<?php bloginfo( 'name' ); ?> &mdash; <span><?php echo get_bloginfo( 'description' ); ?></span>
			 <? if ( is_single() ) :
        echo '</h3>';
    else :
        echo '</h1>';
    endif; ?>
			
		</div><!-- /brand -->
	
		<nav role="navigation" class="site-navigation main-navigation">
			<?php wp_nav_menu( array( 'theme_location' => 'primary' ) ); ?>
		</nav><!-- .site-navigation .main-navigation -->
		
		<div class="clear"></div>
	</div><!--/container -->
		
</header><!-- #masthead .site-header -->

<div class="container">

	<div id="primary">
		<div id="content" role="main">


<?php
	/*-----------------------------------------------------------------------------------*/
	/* Start Home loop
	/*-----------------------------------------------------------------------------------*/
	
	if( is_home() || is_archive() ) {
	
?>
			<?php if ( have_posts() ) : ?>

				<?php while ( have_posts() ) : the_post(); ?>

					<div class="post_main">
					
						<h2 class="title">
							<a href="<?php the_permalink(); ?>" title="<?php the_title(); ?>">
								<?php the_title() ?>
							</a>
						</h2>
						
						
						<div class="the-content">
							<?php the_content( 'продолжение...' ); ?>
							
							<?php wp_link_pages(); ?>
						</div><!-- the-content -->
						
						<div class="meta clearfix">
							
							<div class="arrow_top"><a href="/">	<?php echo file_get_contents($icon_arrow);  ?> </a></div>
							<div class="date"><?php echo get_the_date('j-m-Y'); ?></div>
							<div class="comment"> <?php if( comments_open() ) : ?>
								<span class="comments-link"><?php comments_popup_link('комментариев нет', '1 комментарий', '% комментариев', 'с_link'); ?>
								</span>
							<?php endif; ?>
						    </div>
						</div><!-- Meta -->
						
					</div>

				<?php endwhile; ?>
				
				<!-- pagintation -->
				<?php //the_posts_pagination(); ?>
				<div id="pagination" class="clearfix">
					<div class="past-page"><?php previous_posts_link( 'Вперед  <svg width="42" height="12" viewBox="0 0 42 16" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M33.0615 1.6665L41 8.33317L33.0615 14.9998M1 8.33318H40.2349" stroke="#BB484A" stroke-width="2" stroke-miterlimit="10" stroke-linecap="round" stroke-linejoin="round"/>
</svg>' ); ?></div>
					<div class="next-page"><?php next_posts_link( ' <svg width="42" height="12" viewBox="0 0 42 16" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M8.93853 1.6665L1 8.33317L8.93853 14.9998M41 8.33318H1.76508" stroke="#BB484A" stroke-width="2" stroke-miterlimit="10" stroke-linecap="round" stroke-linejoin="round"/>
</svg>  Назад' ); ?></div>
				</div><!-- pagination -->


			<?php else : ?>
				
				<article class="post error">
					<h1 class="404">Nothing posted yet</h1>
				</article>

			<?php endif; ?>

		
	<?php } //end is_home(); ?>

<?php
	/*-----------------------------------------------------------------------------------*/
	/* Start Single loop
	/*-----------------------------------------------------------------------------------*/
	
	if( is_single() ) {
?>


			<?php if ( have_posts() ) : ?>

				<?php while ( have_posts() ) : the_post(); ?>

					<article class="post">
					
						<h1 class="title"><?php the_title() ?></h1>
						
						
						<div class="the-content">
							<?php the_content( 'Continue...' ); ?>
							
							<?php wp_link_pages(); ?>
						</div><!-- the-content -->
						
						<div class="meta clearfix">
							
							<div class="arrow_top"><a href="/">	<?php echo file_get_contents($icon_arrow);  ?> </a></div>
							<div class="date"><?php echo get_the_date('j-m-Y'); ?></div>
							<div class="comment"> <?php if( comments_open() ) : ?>
								<span class="comments-link"><?php comments_popup_link('комментариев нет', '1 комментарий', '% комментариев', 'с_link'); ?>
								</span>
							<?php endif; ?>
						    </div>
						</div><!-- Meta -->					
						
					</article>

				<?php endwhile; ?>
				
				<?php
					// If comments are open or we have at least one comment, load up the comment template
					if ( comments_open() || '0' != get_comments_number() )
						comments_template( '', true );
				?>


			<?php else : ?>
				
				<article class="post error">
					<h1 class="404">Что то пошло не так ...</h1>
				</article>

			<?php endif; ?>


	<?php } //end is_single(); ?>
	
<?php
	/*-----------------------------------------------------------------------------------*/
	/* Start Page loop
	/*-----------------------------------------------------------------------------------*/
	
	if( is_page()) {
?>

			<?php if ( have_posts() ) : ?>

				<?php while ( have_posts() ) : the_post(); ?>

					<article class="post">
					
						<h1 class="title"><?php the_title() ?></h1>
						
						<div class="the-content">
							<?php the_content(); ?>
							
							<?php wp_link_pages(); ?>
						</div><!-- the-content -->
						
					</article>

				<?php endwhile; ?>

			<?php else : ?>
				
				<article class="post error">
					<h1 class="404">Что то пошло не так ...</h1>
				</article>

			<?php endif; ?>

	<?php } // end is_page(); ?>

		</div><!-- #content .site-content -->
	</div><!-- #primary .content-area -->

</div><!-- / container-->

<?php
	/*-----------------------------------------------------------------------------------*/
	/* Start Footer
	/*-----------------------------------------------------------------------------------*/
?>

<footer class="site-footer" role="contentinfo"> 
	<div class="site-info container">
	    <?php echo file_get_contents($trud_svg); ?>
	    <hr>
	    <span>© Охрана труда. ООО "Лига Качества", Киров, <?php echo date("Y"); ?></span> | <a href="/privacy-policy/">Политика конфиденциальности</a> <!-- | <a href="mailto:info@trudnet,ru">info@trudnet,ru</a> --> 
		
	</div><!-- .site-info -->
</footer><!-- #colophon .site-footer -->

<?php wp_footer(); ?>
<!-- Yandex.Metrika counter -->
<script type="text/javascript" >
   (function(m,e,t,r,i,k,a){m[i]=m[i]||function(){(m[i].a=m[i].a||[]).push(arguments)};
   m[i].l=1*new Date();
   for (var j = 0; j < document.scripts.length; j++) {if (document.scripts[j].src === r) { return; }}
   k=e.createElement(t),a=e.getElementsByTagName(t)[0],k.async=1,k.src=r,a.parentNode.insertBefore(k,a)})
   (window, document, "script", "https://mc.yandex.ru/metrika/tag.js", "ym");

   ym(97005514, "init", {
        clickmap:true,
        trackLinks:true,
        accurateTrackBounce:true,
        webvisor:true
   });
</script>
<noscript><div><img src="https://mc.yandex.ru/watch/97005514" style="position:absolute; left:-9999px;" alt="" /></div></noscript>
<!-- /Yandex.Metrika counter -->

</body>
</html>
