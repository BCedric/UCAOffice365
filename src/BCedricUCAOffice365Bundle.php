<?php

namespace BCedric\UCAOffice365;

use BCedric\UCAOffice365\Service\UCAOffice365;
use Symfony\Component\Config\Definition\Configurator\DefinitionConfigurator;
use Symfony\Component\DependencyInjection\ContainerBuilder;
use Symfony\Component\DependencyInjection\Loader\Configurator\ContainerConfigurator;
use Symfony\Component\HttpKernel\Bundle\AbstractBundle;



class BCedricUCAOffice365Bundle extends AbstractBundle
{

    public function configure(DefinitionConfigurator $definition): void
    {
        $definition->rootNode()
            ->children()
            ->scalarNode('url')->end()
            ->scalarNode('login')->end()
            ->scalarNode('password')->end()
            ->end();
    }

    public function loadExtension(array $config, ContainerConfigurator $container, ContainerBuilder $builder): void
    {
        $container->import('../config/services.yaml');

        $container->services()->set(UCAOffice365::class)
            ->public();

        $builder->autowire(UCAOffice365::class)
            ->setArgument('$url', $config['url'])
            ->setArgument('$login', $config['login'])
            ->setArgument('$password', $config['password'])
        ;
    }
}
