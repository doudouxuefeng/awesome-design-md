/* ============================================================
   睿博检测官网 — 共享交互脚本
   仅基础交互：导航状态 / 移动端菜单 / 表单校验 / 基础淡入
   ============================================================ */

(function(){
  'use strict';

  /* ---------- 导航栏滚动状态 ---------- */
  var header = document.querySelector('.header');
  if (header){
    window.addEventListener('scroll', function(){
      if (window.scrollY > 10) header.style.boxShadow = '0 1px 6px rgba(0,0,0,.06)';
      else header.style.boxShadow = 'none';
    });
  }

  /* ---------- 移动端菜单 ---------- */
  var hamburger = document.getElementById('hamburger');
  var mobileMenu = document.getElementById('mobileMenu');
  if (hamburger && mobileMenu){
    hamburger.addEventListener('click', function(){
      hamburger.classList.toggle('open');
      mobileMenu.classList.toggle('open');
    });
    mobileMenu.querySelectorAll('a').forEach(function(a){
      a.addEventListener('click', function(){
        hamburger.classList.remove('open');
        mobileMenu.classList.remove('open');
      });
    });
  }

  /* ---------- 基础淡入（克制，无滚动视差） ---------- */
  var reveals = document.querySelectorAll('.reveal');
  if (reveals.length && 'IntersectionObserver' in window){
    var io = new IntersectionObserver(function(entries){
      entries.forEach(function(e){
        if (e.isIntersecting){
          e.target.classList.add('in');
          io.unobserve(e.target);
        }
      });
    }, { threshold: 0.1 });
    reveals.forEach(function(el){ io.observe(el); });
  } else {
    reveals.forEach(function(el){ el.classList.add('in'); });
  }

  /* ---------- 表单校验与提交（mock） ---------- */
  var forms = document.querySelectorAll('[data-form]');
  forms.forEach(function(form){
    form.addEventListener('submit', function(e){
      e.preventDefault();
      var name = form.querySelector('[name="name"]');
      var phone = form.querySelector('[name="phone"]');
      var industry = form.querySelector('[name="industry"]');

      if (name && !name.value.trim()){
        alert('请填写您的姓名'); name.focus(); return;
      }
      if (phone && !/^1[3-9]\d{9}$/.test(phone.value.trim())){
        alert('请输入正确的 11 位手机号码'); phone.focus(); return;
      }
      if (industry && !industry.value){
        alert('请选择所属行业'); industry.focus(); return;
      }

      form.style.display = 'none';
      var success = form.parentElement.querySelector('.form-success');
      if (success) success.classList.add('show');
    });
  });

})();
